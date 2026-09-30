// 「처음 AI 에이전트를 만드는 개발팀에게」 배포 슬라이드 — 콘텐츠 전략 v2 기준 골격
// 실행: node deck/build_deck.js  →  deck/first-ai-agent-team.pptx
//
// 원문(01_first-ai-agent-team-checklist_v0.2.md)이 아직 없어 본문 중 원문에서 와야 하는
// 자리는 점선 상자 `〔원문 입력〕`으로 비워 두었다. 전략 문서에 명시된 제목·헤더밴드 결론·
// 구조·항목명·수치(13명·4명, 10가지, 5단계, 6가지, 8개)만 채웠다.
// 각주의 절 번호는 SOURCE 맵에서 채운다.

const pptxgen = require("pptxgenjs");
const path = require("path");

// 원문 절 번호: 원문 확보 후 "§3.2" 형태로 교체
const SOURCE = {
  1: "§?", 2: "§?", 3: "§?", 4: "§?", 5: "§?", 6: "§?", 7: "§?",
  8: "§?", 9: "§?", 10: "§?", 11: "§?", 12: "§?", 13: "§?", 14: "§?",
};
const SOURCE_DOC = "01_first-ai-agent-team-checklist_v0.2.md";

// ---------- Design tokens (Design System v3 자산-없음 폴백) ----------
const W = 13.333, H = 7.5;
const SB = 2.35;                 // 사이드바 폭
const BX = 2.75, BW = W - BX - 0.45;
const BY = 1.95, BB = 6.7;       // 본문 영역 상·하단
const FONT = "Malgun Gothic";
const C = {
  ink: "1B2436", ink2: "2C3A52", band: "F0F3F7", text: "1F2933", muted: "6B7785",
  line: "D5DBE3", accent: "E4572E", accentSoft: "FCE9E3", teal: "1F7A8C", tealSoft: "E1F0F3",
  phLine: "9AA5B1", phFill: "FAFBFC", white: "FFFFFF", side: "8A96A8",
};

const ACTS = [
  { id: "Ⅰ", name: "왜 필요한가", range: [1, 3] },
  { id: "Ⅱ", name: "시작 전 합의", range: [4, 6] },
  { id: "Ⅲ", name: "설계·구현 통제", range: [7, 9] },
  { id: "Ⅳ", name: "검증·배포·역할", range: [10, 14] },
];
const TOTAL = 17;

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.title = "처음 AI 에이전트를 만드는 개발팀에게";
pres.theme = { headFontFace: FONT, bodyFontFace: FONT };

// ---------- helpers ----------
const T = (s, text, o) => s.addText(text, { fontFace: FONT, color: C.text, fontSize: 12, isTextBox: true, margin: 0, valign: "top", ...o });
const box = (s, x, y, w, h, o = {}) =>
  s.addShape(o.rounded === false ? pres.shapes.RECTANGLE : pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h, fill: { color: o.fill || C.white }, line: o.line ? { color: o.line, width: o.lw || 0.75, dashType: o.dash } : { type: "none" },
    ...(o.rounded === false ? {} : { rectRadius: o.r ?? 0.08 }),
  });

// 원문에서 채울 자리
function ph(s, x, y, w, h, label, o = {}) {
  box(s, x, y, w, h, { fill: C.phFill, line: C.phLine, dash: "dash", r: 0.06 });
  T(s, [
    { text: "〔원문 입력〕", options: { color: C.accent, bold: true, fontSize: o.fs || 9, breakLine: !!label } },
    ...(label ? [{ text: label, options: { color: C.muted, fontSize: o.fs || 9 } }] : []),
  ], { x: x + 0.1, y: y + 0.07, w: w - 0.2, h: h - 0.12, valign: o.valign || "top" });
}

function pill(s, x, y, w, h, text, o = {}) {
  box(s, x, y, w, h, { fill: o.fill || C.ink, r: h / 2 });
  T(s, text, { x, y, w, h, align: "center", valign: "middle", fontSize: o.fs || 11, bold: true, color: o.color || C.white });
}

function shell(s, n, title, conclusion, notes) {
  s.background = { color: C.white };
  const act = ACTS.find(a => n >= a.range[0] && n <= a.range[1]);
  // sidebar
  box(s, 0, 0, SB, H, { fill: C.ink, rounded: false });
  T(s, "처음 AI 에이전트를\n만드는 개발팀에게", { x: 0.35, y: 0.4, w: SB - 0.6, h: 0.7, fontSize: 11, bold: true, color: C.white });
  ACTS.forEach((a, i) => {
    const y = 1.7 + i * 0.95, on = a === act;
    if (on) box(s, 0.2, y - 0.1, SB - 0.4, 0.8, { fill: C.ink2 });
    T(s, a.id, { x: 0.35, y, w: 0.35, h: 0.35, fontSize: 14, bold: true, color: on ? C.accent : C.side });
    T(s, a.name, { x: 0.72, y, w: SB - 1.0, h: 0.32, fontSize: 11, bold: on, color: on ? C.white : C.side });
    T(s, `${a.range[0]}–${a.range[1]}장`, { x: 0.72, y: y + 0.32, w: SB - 1.0, h: 0.28, fontSize: 9, color: C.side });
  });
  T(s, `${String(n + 2).padStart(2, "0")} / ${TOTAL}`, { x: 0.35, y: 6.85, w: 1.5, h: 0.3, fontSize: 9, color: C.side });
  // header band
  box(s, SB, 0, W - SB, 1.6, { fill: C.band, rounded: false });
  T(s, `${String(n).padStart(2, "0")}   ${title}`, { x: BX, y: 0.32, w: BW, h: 0.34, fontSize: 12, bold: true, color: C.accent });
  T(s, conclusion, { x: BX, y: 0.7, w: BW, h: 0.75, fontSize: 21, bold: true, color: C.ink });
  // footnote
  T(s, `출처 · 원문 ${SOURCE[n]}  ${SOURCE_DOC}`, { x: BX, y: 6.98, w: BW, h: 0.3, fontSize: 9, color: C.muted });
  if (notes) s.addNotes(notes);
}

function tableGrid(s, x, y, w, rows, cols, o = {}) {
  // cols: [{label, w}] ; rows: [[cell,...]] where cell is string or {ph:label} or {text, bold}
  const hh = o.headH || 0.4, rh = o.rowH || 0.5;
  let cx = x;
  cols.forEach(c => {
    box(s, cx, y, c.w - 0.04, hh, { fill: o.headFill || C.ink, r: 0.04 });
    T(s, c.label, { x: cx + 0.1, y, w: c.w - 0.2, h: hh, fontSize: 10.5, bold: true, color: C.white, valign: "middle" });
    cx += c.w;
  });
  rows.forEach((r, ri) => {
    let cx2 = x; const ry = y + hh + 0.06 + ri * (rh + 0.06);
    r.forEach((cell, ci) => {
      const cw = cols[ci].w - 0.04;
      if (cell && cell.ph !== undefined) ph(s, cx2, ry, cw, rh, cell.ph);
      else {
        box(s, cx2, ry, cw, rh, { fill: cell && cell.fill || C.band, r: 0.04 });
        T(s, cell && cell.text !== undefined ? cell.text : cell, { x: cx2 + 0.1, y: ry, w: cw - 0.2, h: rh, fontSize: 10.5, bold: !!(cell && cell.bold), color: cell && cell.color || C.text, valign: "middle" });
      }
      cx2 += cols[ci].w;
    });
  });
}

// ---------- 표지 ----------
{
  const s = pres.addSlide();
  s.background = { color: C.ink };
  T(s, "실무 안내 · 단계별 체크리스트", { x: 0.9, y: 1.3, w: 8, h: 0.4, fontSize: 14, bold: true, color: C.accent });
  T(s, "처음 AI 에이전트를\n만드는 개발팀에게", { x: 0.9, y: 1.9, w: 11, h: 2.2, fontSize: 44, bold: true, color: C.white });
  T(s, "처음 AI 에이전트를 만드는 팀이 덜 헤매도록,\n단계별로 꼭 확인할 질문과 기준을 정리했습니다.", { x: 0.9, y: 4.3, w: 10, h: 1.0, fontSize: 18, color: "CBD3DE" });
  T(s, "대상  개발자 · PM · 프로젝트 리더", { x: 0.9, y: 6.4, w: 8, h: 0.4, fontSize: 12, color: C.side });
  s.addNotes("표지. 로고·사진 자산 없이 완결되는 Design System v3 자산-없음 폴백.");
}

// ---------- 목차 ----------
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  box(s, 0, 0, W, 1.6, { fill: C.band, rounded: false });
  T(s, "목차", { x: 0.7, y: 0.32, w: 6, h: 0.34, fontSize: 12, bold: true, color: C.accent });
  T(s, "4막 14장: 왜 필요한가에서 출시 직전 점검까지", { x: 0.7, y: 0.7, w: 12, h: 0.7, fontSize: 24, bold: true, color: C.ink });
  const titles = {
    1: "AI 에이전트 프로젝트, 어디가 다른가", 2: "현장에서 반복된 네 가지 어려움", 3: "이 자료가 하는 일과 하지 않는 일",
    4: "착수 체크리스트", 5: "요구사항이 바뀌는 프로젝트를 다루는 법", 6: "구현 전 반드시 답할 10가지 질문",
    7: "설계 체크리스트", 8: "도구 위험 사다리", 9: "구현 체크리스트",
    10: "검증 체크리스트", 11: "배포 체크리스트", 12: "설명 가능성", 13: "PM·리더가 먼저 확보해야 하는 여섯 가지", 14: "출시 직전 점검과 다음 행동",
  };
  const cw = (W - 1.4 - 0.3 * 3) / 4;
  ACTS.forEach((a, i) => {
    const x = 0.7 + i * (cw + 0.3), y = 2.0;
    box(s, x, y, cw, 4.7, { fill: i % 2 ? C.white : C.band, line: C.line });
    T(s, a.id, { x: x + 0.25, y: y + 0.25, w: 1, h: 0.6, fontSize: 30, bold: true, color: C.accent });
    T(s, a.name, { x: x + 0.25, y: y + 0.95, w: cw - 0.5, h: 0.4, fontSize: 15, bold: true, color: C.ink });
    const items = [];
    for (let n = a.range[0]; n <= a.range[1]; n++) items.push({ text: `${n}  ${titles[n]}`, options: { breakLine: n < a.range[1], paraSpaceAfter: 8 } });
    T(s, items, { x: x + 0.25, y: y + 1.55, w: cw - 0.5, h: 3.0, fontSize: 11, color: C.text });
  });
  s.addNotes("목차. 사이드바 챕터 = 4막.");
}

// ---------- 1 ----------
{
  const s = pres.addSlide();
  shell(s, 1, "AI 에이전트 프로젝트, 어디가 다른가", "SI 기본기는 그대로 핵심이며, 불확실성 확인 루프만 더해집니다",
    "기존 SI ↔ AI 에이전트 5단계 대조. 좌열(SI 기본기)은 유지된다는 점을 강조. 교차 원칙: SI 기본기(1·3·7장).");
  const lw = 1.5, cw = (BW - lw - 0.2) / 2;
  const x1 = BX + lw + 0.1, x2 = x1 + cw + 0.1;
  pill(s, x1, BY, cw, 0.45, "기존 SI · 그대로 유지", { fill: C.ink });
  pill(s, x2, BY, cw, 0.45, "AI 에이전트 · 불확실성 확인 루프 추가", { fill: C.accent });
  T(s, "SI 기본기", { x: x1, y: BY + 0.5, w: cw, h: 0.3, fontSize: 9, color: C.muted, align: "center" });
  for (let i = 0; i < 5; i++) {
    const y = BY + 0.85 + i * 0.78;
    ph(s, BX, y, lw, 0.68, `단계 ${i + 1}`);
    ph(s, x1, y, cw, 0.68, "기존 SI에서 하는 일");
    box(s, x2, y, cw, 0.68, { fill: C.accentSoft, r: 0.06 });
    T(s, [{ text: "+ ", options: { bold: true, color: C.accent } }, { text: "〔원문 입력〕 더해지는 불확실성 확인", options: { color: C.muted, fontSize: 9 } }],
      { x: x2 + 0.12, y, w: cw - 0.24, h: 0.68, valign: "middle", fontSize: 12 });
  }
}

// ---------- 2 ----------
{
  const s = pres.addSlide();
  shell(s, 2, "현장에서 반복된 네 가지 어려움", "팀이 헤매는 지점은 네 가지 신호로 반복됩니다",
    "조사 범위 띠 + 신호 4 + 익명 인용 4. 수치는 원문에 있는 설문 13명·4명만 사용. 교차 원칙: 기록된 가정과 변경 관리(현장 인용).");
  box(s, BX, BY, BW, 0.7, { fill: C.band });
  T(s, "조사 범위", { x: BX + 0.2, y: BY, w: 1.3, h: 0.7, fontSize: 11, bold: true, color: C.muted, valign: "middle" });
  T(s, [{ text: "13명", options: { fontSize: 24, bold: true, color: C.ink } }, { text: "  설문", options: { fontSize: 11, color: C.muted } }],
    { x: BX + 1.5, y: BY, w: 2.2, h: 0.7, valign: "middle" });
  T(s, [{ text: "4명", options: { fontSize: 24, bold: true, color: C.ink } }, { text: "  〔원문 입력〕 구분", options: { fontSize: 11, color: C.muted } }],
    { x: BX + 3.7, y: BY, w: 3.0, h: 0.7, valign: "middle" });
  const cw = (BW - 0.3 * 3) / 4;
  for (let i = 0; i < 4; i++) {
    const x = BX + i * (cw + 0.3), y = BY + 1.0;
    box(s, x, y, cw, 3.7, { fill: C.white, line: C.line });
    T(s, `신호 ${i + 1}`, { x: x + 0.2, y: y + 0.2, w: cw - 0.4, h: 0.4, fontSize: 18, bold: true, color: C.accent });
    ph(s, x + 0.2, y + 0.7, cw - 0.4, 0.95, "어려움 신호");
    T(s, "“", { x: x + 0.2, y: y + 1.8, w: 0.5, h: 0.6, fontSize: 36, bold: true, color: C.line });
    ph(s, x + 0.2, y + 2.35, cw - 0.4, 1.15, "익명 인용");
  }
}

// ---------- 3 ----------
{
  const s = pres.addSlide();
  shell(s, 3, "이 자료가 하는 일과 하지 않는 일", "질문·증빙·책임자는 제공하고, 방법론·수치는 강제하지 않습니다",
    "2열 대조 + 단계별 사용 지도. '하지 않는 일'의 마지막 행에 SI 기본기(교차 원칙).");
  const cw = (BW - 0.3) / 2;
  const cols = [
    { head: "하는 일", fill: C.teal, items: ["단계별로 확인할 질문", "확인했다는 증빙", "확인 책임자"] },
    { head: "하지 않는 일", fill: C.ink, items: ["특정 방법론 강제", "목표 수치 강제", "SI 기본기를 대체"] },
  ];
  cols.forEach((c, i) => {
    const x = BX + i * (cw + 0.3);
    pill(s, x, BY, cw, 0.45, c.head, { fill: c.fill });
    c.items.forEach((it, j) => {
      const y = BY + 0.6 + j * 0.62;
      const last = i === 1 && j === 2;
      box(s, x, y, cw, 0.52, { fill: last ? C.accentSoft : C.band, r: 0.06 });
      T(s, it, { x: x + 0.2, y, w: cw - 0.4, h: 0.52, fontSize: 13, bold: last, valign: "middle", color: last ? C.accent : C.text });
    });
  });
  const my = BY + 2.65;
  T(s, "단계별 사용 지도", { x: BX, y: my, w: 4, h: 0.35, fontSize: 13, bold: true, color: C.ink });
  const steps = [["Ⅰ 왜 필요한가", "1–3장"], ["Ⅱ 시작 전 합의", "4–6장"], ["Ⅲ 설계·구현 통제", "7–9장"], ["Ⅳ 검증·배포·역할", "10–14장"]];
  const sw = (BW - 0.25 * 3) / 4;
  steps.forEach(([a, r], i) => {
    const x = BX + i * (sw + 0.25), y = my + 0.5;
    s.addShape(pres.shapes.CHEVRON, { x, y, w: sw, h: 0.75, fill: { color: i === 0 ? C.line : C.ink2 }, line: { type: "none" } });
    T(s, a, { x: x + 0.35, y, w: sw - 0.7, h: 0.75, fontSize: 11, bold: true, color: i === 0 ? C.ink : C.white, valign: "middle", align: "center" });
    ph(s, x, y + 0.9, sw, 0.75, `${r} 사용 시점`);
  });
}

// ---------- 4 ----------
{
  const s = pres.addSlide();
  shell(s, 4, "착수 체크리스트", "코딩 전에 대표 업무·실제 데이터·권한·사용 가능 기준을 문서로 남깁니다",
    "세로 Pill + 4열 표(항목·확인 질문·증빙·책임자) + 함정 3. 부록 상세 질문은 표 셀에 흡수.");
  const items = ["대표 업무", "실제 데이터", "권한", "사용 가능 기준"];
  const tw = BW - 3.1;
  const cols = [{ label: "항목", w: 1.8 }, { label: "확인 질문", w: (tw - 1.8) * 0.45 }, { label: "증빙", w: (tw - 1.8) * 0.33 }, { label: "책임자", w: (tw - 1.8) * 0.22 }];
  const hh = 0.4, rh = 0.9;
  tableGrid(s, BX, BY, tw, items.map(() => [null, { ph: "" }, { ph: "" }, { ph: "" }]), cols, { headH: hh, rowH: rh });
  items.forEach((it, i) => pill(s, BX, BY + hh + 0.06 + i * (rh + 0.06) + 0.2, 1.72, 0.5, it, { fs: 11 }));
  const fx = BX + tw + 0.3, fw = BW - tw - 0.3;
  T(s, "흔한 함정", { x: fx, y: BY, w: fw, h: 0.4, fontSize: 13, bold: true, color: C.accent, valign: "middle" });
  for (let i = 0; i < 3; i++) {
    const y = BY + 0.5 + i * 1.42;
    box(s, fx, y, fw, 1.3, { fill: C.accentSoft });
    T(s, `함정 ${i + 1}`, { x: fx + 0.15, y: y + 0.1, w: fw - 0.3, h: 0.3, fontSize: 11, bold: true, color: C.accent });
    ph(s, fx + 0.15, y + 0.45, fw - 0.3, 0.72, "");
  }
}

// ---------- 5 ----------
{
  const s = pres.addSlide();
  shell(s, 5, "요구사항이 바뀌는 프로젝트를 다루는 법", "변경 자체가 아니라, 변경 이유와 영향을 남기지 않는 것이 실패입니다",
    "4단계 흐름 + 가정 기록 항목 + PM 합의. 교차 원칙: 기록된 가정과 변경 관리.");
  const sw = (BW - 0.2 * 3) / 4;
  for (let i = 0; i < 4; i++) {
    const x = BX + i * (sw + 0.2);
    s.addShape(pres.shapes.CHEVRON, { x, y: BY, w: sw, h: 0.8, fill: { color: i === 3 ? C.accent : C.ink2 }, line: { type: "none" } });
    T(s, `${i + 1}단계`, { x: x + 0.35, y: BY, w: sw - 0.7, h: 0.8, fontSize: 14, bold: true, color: C.white, valign: "middle", align: "center" });
    ph(s, x, BY + 0.95, sw, 1.1, "단계 설명");
  }
  const y2 = BY + 2.35, hw = (BW - 0.3) / 2;
  box(s, BX, y2, hw, 2.35, { fill: C.band });
  T(s, "가정 기록 항목", { x: BX + 0.2, y: y2 + 0.15, w: hw - 0.4, h: 0.35, fontSize: 13, bold: true, color: C.ink });
  T(s, "변경 이유 · 영향 범위를 함께 남긴다", { x: BX + 0.2, y: y2 + 0.5, w: hw - 0.4, h: 0.3, fontSize: 10, color: C.muted });
  ph(s, BX + 0.2, y2 + 0.9, hw - 0.4, 1.3, "기록 항목 목록");
  const x3 = BX + hw + 0.3;
  box(s, x3, y2, hw, 2.35, { fill: C.tealSoft });
  T(s, "PM 합의", { x: x3 + 0.2, y: y2 + 0.15, w: hw - 0.4, h: 0.35, fontSize: 13, bold: true, color: C.teal });
  T(s, "누가 · 언제 변경을 승인하는가", { x: x3 + 0.2, y: y2 + 0.5, w: hw - 0.4, h: 0.3, fontSize: 10, color: C.muted });
  ph(s, x3 + 0.2, y2 + 0.9, hw - 0.4, 1.3, "합의 대상과 시점");
}

// ---------- 6 ----------
{
  const s = pres.addSlide();
  shell(s, 6, "구현 전 반드시 답할 10가지 질문", "답하기 어렵다면, 구현보다 답을 찾는 일이 먼저입니다",
    "3묶음 체크리스트, 원문 번호 유지. 묶음 구분(4·3·3)은 원문 기준으로 조정.");
  const groups = [[1, 2, 3, 4], [5, 6, 7], [8, 9, 10]];
  const cw = (BW - 0.3 * 2) / 3;
  groups.forEach((g, i) => {
    const x = BX + i * (cw + 0.3);
    box(s, x, BY, cw, 4.7, { fill: C.band });
    T(s, `묶음 ${String.fromCharCode(65 + i)}`, { x: x + 0.2, y: BY + 0.15, w: 1.3, h: 0.35, fontSize: 13, bold: true, color: C.accent });
    ph(s, x + 1.3, BY + 0.12, cw - 1.5, 0.4, "묶음 이름");
    g.forEach((q, j) => {
      const y = BY + 0.7 + j * 0.97;
      box(s, x + 0.2, y + 0.08, 0.32, 0.32, { fill: C.white, line: C.ink, lw: 1.25, r: 0.04 });
      T(s, `Q${q}`, { x: x + 0.62, y: y + 0.05, w: 0.6, h: 0.38, fontSize: 13, bold: true, color: C.ink, valign: "middle" });
      ph(s, x + 1.2, y, cw - 1.4, 0.85, "");
    });
  });
}

// ---------- 7 ----------
{
  const s = pres.addSlide();
  shell(s, 7, "설계 체크리스트", "규칙·검증·권한은 코드가, 판단·생성은 모델이 맡고, 외부 입력은 데이터로 다룹니다",
    "신뢰 경계 도식 + 6행 표. 교차 원칙: SI 기본기(코드 영역이 설계의 축), 설명 가능성('설명 설계' 행).");
  const dw = 4.3, dx = BX, dy = BY;
  // 외부 입력
  box(s, dx, dy, dw, 0.8, { fill: C.band, line: C.phLine, dash: "dash" });
  T(s, [{ text: "외부 입력  ", options: { bold: true, color: C.ink } }, { text: "= 데이터 (지시로 취급하지 않음)", options: { color: C.muted, fontSize: 10 } }],
    { x: dx + 0.2, y: dy, w: dw - 0.4, h: 0.8, fontSize: 12, valign: "middle" });
  s.addShape(pres.shapes.DOWN_ARROW, { x: dx + dw / 2 - 0.2, y: dy + 0.88, w: 0.4, h: 0.4, fill: { color: C.phLine }, line: { type: "none" } });
  T(s, "신뢰 경계", { x: dx + dw / 2 + 0.3, y: dy + 0.9, w: 1.5, h: 0.35, fontSize: 9, bold: true, color: C.accent, valign: "middle" });
  // 코드 영역
  const cy = dy + 1.38;
  box(s, dx, cy, dw, 3.3, { fill: C.ink });
  T(s, "코드", { x: dx + 0.25, y: cy + 0.2, w: 1.5, h: 0.4, fontSize: 16, bold: true, color: C.white });
  pill(s, dx + dw - 1.35, cy + 0.22, 1.1, 0.34, "SI 기본기", { fill: C.accent, fs: 9 });
  T(s, "규칙 · 검증 · 권한", { x: dx + 0.25, y: cy + 0.65, w: dw - 0.5, h: 0.35, fontSize: 12, color: "CBD3DE" });
  box(s, dx + 0.45, cy + 1.25, dw - 0.9, 1.75, { fill: C.teal });
  T(s, "모델", { x: dx + 0.7, y: cy + 1.45, w: 2, h: 0.4, fontSize: 16, bold: true, color: C.white });
  T(s, "판단 · 생성\n코드가 경계 안에서 호출하고 결과를 검증", { x: dx + 0.7, y: cy + 1.9, w: dw - 1.4, h: 0.9, fontSize: 11, color: C.white });
  // 6행 표
  const tx = BX + dw + 0.35, tw = BW - dw - 0.35;
  const rows = [1, 2, 3, 4, 5].map(() => [{ ph: "항목" }, { ph: "확인 기준" }]);
  rows.push([{ text: "설명 설계", bold: true, color: C.teal, fill: C.tealSoft }, { ph: "누가 무엇을 설명받는가" }]);
  tableGrid(s, tx, BY, tw, rows, [{ label: "설계 항목", w: 1.9 }, { label: "확인 기준", w: tw - 1.9 }], { rowH: 0.64 });
}

// ---------- 8 ----------
{
  const s = pres.addSlide();
  shell(s, 8, "도구 위험 사다리", "도구는 위험도에 따라 통제하고, 승인이 과다하면 자율 범위를 다시 봅니다",
    "5단 계단 + 예시·통제 행. 교차 원칙: 실행 통제(에이전트 고유).");
  const lw = 1.0, sw = (BW - lw - 0.1 * 4) / 5;
  const baseY = BY + 2.3;
  T(s, "예시", { x: BX, y: baseY + 0.1, w: lw, h: 0.85, fontSize: 12, bold: true, color: C.muted, valign: "middle" });
  T(s, "통제", { x: BX, y: baseY + 1.05, w: lw, h: 0.85, fontSize: 12, bold: true, color: C.muted, valign: "middle" });
  for (let i = 0; i < 5; i++) {
    const x = BX + lw + i * (sw + 0.1), h = 0.55 + i * 0.4;
    const shade = ["C9D2DD", "A7B4C4", "7F8FA6", "4A5C78", C.accent][i];
    box(s, x, baseY - h, sw, h, { fill: shade, r: 0.05, rounded: false });
    T(s, `${i + 1}단계`, { x: x + 0.1, y: baseY - h + 0.08, w: sw - 0.2, h: 0.35, fontSize: 13, bold: true, color: i < 2 ? C.ink : C.white });
    ph(s, x, baseY + 0.1, sw, 0.85, "도구 예시");
    ph(s, x, baseY + 1.05, sw, 0.85, "통제 방식");
  }
  s.addShape(pres.shapes.RIGHT_ARROW, { x: BX + lw, y: BY - 0.02, w: 2.6, h: 0.28, fill: { color: C.line }, line: { type: "none" } });
  T(s, "위험도 증가", { x: BX + lw + 2.7, y: BY - 0.04, w: 1.5, h: 0.3, fontSize: 10, bold: true, color: C.muted, valign: "middle" });
  const cy = baseY + 2.02;
  box(s, BX, cy, BW, 0.4, { fill: C.accentSoft });
  T(s, [{ text: "승인이 과다하다면  →  ", options: { bold: true, color: C.accent } }, { text: "통제를 늘리기 전에 에이전트의 자율 범위를 재검토", options: { color: C.text } }],
    { x: BX + 0.2, y: cy, w: BW - 0.4, h: 0.4, fontSize: 11, valign: "middle" });
}

// ---------- 9 ----------
{
  const s = pres.addSlide();
  shell(s, 9, "구현 체크리스트", "중단·중복 방지·이관·비상 중단을 기능과 함께 구현합니다",
    "실행 타임라인 + 통제 플래그 + 5행 표. 교차 원칙: 실행 통제(타임라인 플래그).");
  const ly = BY + 0.95;
  T(s, "실행 시작", { x: BX, y: ly - 0.55, w: 1.5, h: 0.3, fontSize: 10, color: C.muted });
  T(s, "실행 종료", { x: BX + BW - 1.5, y: ly - 0.55, w: 1.5, h: 0.3, fontSize: 10, color: C.muted, align: "right" });
  s.addShape(pres.shapes.LINE, { x: BX, y: ly, w: BW - 0.15, h: 0, line: { color: C.ink, width: 2.5, endArrowType: "triangle" } });
  const flags = ["중단", "중복 방지", "이관", "비상 중단"];
  flags.forEach((f, i) => {
    const x = BX + 0.7 + i * ((BW - 1.6) / 3);
    s.addShape(pres.shapes.OVAL, { x: x - 0.11, y: ly - 0.11, w: 0.22, h: 0.22, fill: { color: i === 3 ? C.accent : C.teal }, line: { color: C.white, width: 1.5 } });
    pill(s, x - 0.65, ly + 0.22, 1.3, 0.4, f, { fill: i === 3 ? C.accent : C.teal, fs: 11 });
  });
  const rows = [1, 2, 3, 4, 5].map(() => [{ ph: "구현 항목" }, { ph: "완료 기준" }, { ph: "증빙" }]);
  tableGrid(s, BX, ly + 0.95, BW, rows, [{ label: "구현 항목", w: 2.6 }, { label: "완료 기준", w: BW - 2.6 - 2.6 }, { label: "증빙", w: 2.6 }], { rowH: 0.44 });
}

// ---------- 10 ----------
{
  const s = pres.addSlide();
  shell(s, 10, "검증 체크리스트", "결과가 맞아도, 권한 밖에서 행동했다면 실패입니다",
    "4단 판정 흐름 + 6행 표 + 시나리오 띠. 실패 유형·최소 평가 4항목(부록)은 표에 흡수. 교차 원칙: 설명 가능성('설명 가능성' 행).");
  const fw = 3.4;
  for (let i = 0; i < 4; i++) {
    const y = BY + i * 1.05, last = i === 3;
    box(s, BX, y, fw, 0.8, { fill: last ? C.accent : C.ink2 });
    if (last) T(s, [{ text: "판정 4  ", options: { bold: true } }, { text: "권한 범위 안에서 행동했는가?" }], { x: BX + 0.15, y, w: fw - 0.3, h: 0.8, fontSize: 11.5, color: C.white, valign: "middle" });
    else T(s, [{ text: `판정 ${i + 1}  `, options: { bold: true, color: C.white } }, { text: "〔원문 입력〕", options: { color: "CBD3DE", fontSize: 10 } }], { x: BX + 0.15, y, w: fw - 0.3, h: 0.8, fontSize: 11.5, valign: "middle" });
    if (i < 3) s.addShape(pres.shapes.DOWN_ARROW, { x: BX + fw / 2 - 0.12, y: y + 0.82, w: 0.24, h: 0.2, fill: { color: C.phLine }, line: { type: "none" } });
  }
  T(s, "아니면 → 결과와 무관하게 실패", { x: BX, y: BY + 4.25, w: fw, h: 0.35, fontSize: 10.5, bold: true, color: C.accent });
  const tx = BX + fw + 0.35, tw = BW - fw - 0.35;
  const rows = [1, 2, 3, 4, 5].map(() => [{ ph: "검증 항목" }, { ph: "통과 기준" }]);
  rows.push([{ text: "설명 가능성", bold: true, color: C.teal, fill: C.tealSoft }, { ph: "판단 근거를 확인할 수 있는가" }]);
  tableGrid(s, tx, BY, tw, rows, [{ label: "검증 항목", w: 1.9 }, { label: "통과 기준", w: tw - 1.9 }], { rowH: 0.5 });
  const sy = BY + 3.95;
  box(s, tx, sy, tw, 0.7, { fill: C.band });
  T(s, "시나리오", { x: tx + 0.15, y: sy, w: 1.1, h: 0.7, fontSize: 11, bold: true, color: C.ink, valign: "middle" });
  ph(s, tx + 1.3, sy + 0.1, tw - 1.45, 0.5, "정상 · 예외 · 권한 밖 시나리오", { valign: "middle" });
}

// ---------- 11 ----------
{
  const s = pres.addSlide();
  shell(s, 11, "배포 체크리스트", "Trace·민감정보·지표·드리프트·사용자 안내가 갖춰져야 운영으로 전환합니다",
    "Trace 레이어 타임라인 + 4카드. 재현성 기록(부록)은 각주로 반영.");
  T(s, "Trace 레이어", { x: BX, y: BY, w: 3, h: 0.35, fontSize: 13, bold: true, color: C.ink });
  const lanes = 3, lx = BX + 1.6, lw = BW - 1.6;
  for (let i = 0; i < lanes; i++) {
    const y = BY + 0.5 + i * 0.62;
    ph(s, BX, y, 1.5, 0.5, `레이어 ${i + 1}`);
    box(s, lx, y, lw, 0.5, { fill: C.band, r: 0.04 });
    const segs = [[0.02, 0.3], [0.35, 0.25], [0.64, 0.33]];
    const [a, b] = segs[(i + 1) % 3];
    box(s, lx + lw * a, y + 0.1, lw * b * (0.6 + 0.2 * i), 0.3, { fill: [C.ink2, C.teal, C.accent][i], r: 0.04 });
  }
  T(s, "시간 →", { x: BX + BW - 1.0, y: BY + 2.35, w: 1.0, h: 0.25, fontSize: 9, color: C.muted, align: "right" });
  const cards = ["민감정보", "지표", "드리프트", "사용자 안내"];
  const cw = (BW - 0.3 * 3) / 4, cy = BY + 2.75;
  cards.forEach((c, i) => {
    const x = BX + i * (cw + 0.3);
    box(s, x, cy, cw, 1.95, { fill: C.white, line: C.line });
    T(s, c, { x: x + 0.2, y: cy + 0.15, w: cw - 0.4, h: 0.4, fontSize: 14, bold: true, color: C.ink });
    ph(s, x + 0.2, cy + 0.65, cw - 0.4, 1.1, "운영 전환 조건");
  });
}

// ---------- 12 ----------
{
  const s = pres.addSlide();
  shell(s, 12, "설명 가능성", "대상마다 다른 설명을, 단계마다 다른 증빙을 준비합니다",
    "사용자 화면 ↔ 운영 Trace + 3단 증빙 표. 교차 원칙: 설명 가능성(7·10·12장).");
  const hw = (BW - 1.1) / 2;
  // 사용자 화면
  box(s, BX, BY, hw, 2.3, { fill: C.white, line: C.line });
  box(s, BX, BY, hw, 0.35, { fill: C.band, rounded: false });
  [0, 1, 2].forEach(k => s.addShape(pres.shapes.OVAL, { x: BX + 0.15 + k * 0.2, y: BY + 0.12, w: 0.11, h: 0.11, fill: { color: C.phLine }, line: { type: "none" } }));
  T(s, "사용자 화면", { x: BX + 0.9, y: BY, w: hw - 1.1, h: 0.35, fontSize: 10, bold: true, color: C.ink, valign: "middle" });
  ph(s, BX + 0.2, BY + 0.5, hw - 0.4, 1.6, "사용자에게 보여 줄 설명");
  // 운영 Trace
  const x2 = BX + hw + 1.1;
  box(s, x2, BY, hw, 2.3, { fill: C.ink });
  T(s, "운영 Trace", { x: x2 + 0.2, y: BY + 0.05, w: hw - 0.4, h: 0.3, fontSize: 10, bold: true, color: C.white, valign: "middle" });
  box(s, x2 + 0.2, BY + 0.5, hw - 0.4, 1.6, { fill: C.ink2, r: 0.06 });
  T(s, [{ text: "〔원문 입력〕", options: { color: C.accent, bold: true, breakLine: true } }, { text: "운영자가 확인할 기록", options: { color: "CBD3DE" } }],
    { x: x2 + 0.3, y: BY + 0.57, w: hw - 0.6, h: 1.4, fontSize: 9 });
  s.addShape(pres.shapes.LEFT_RIGHT_ARROW, { x: BX + hw + 0.15, y: BY + 0.95, w: 0.8, h: 0.4, fill: { color: C.accent }, line: { type: "none" } });
  T(s, "대상별로\n다른 설명", { x: BX + hw, y: BY + 1.4, w: 1.1, h: 0.5, fontSize: 9, bold: true, color: C.accent, align: "center" });
  const cols = [{ label: "단계", w: 1.6 }, { label: "필요한 증빙", w: (BW - 1.6) * 0.6 }, { label: "확인 대상", w: (BW - 1.6) * 0.4 }];
  tableGrid(s, BX, BY + 2.6, BW, [1, 2, 3].map(i => [{ ph: `단계 ${i}` }, { ph: "" }, { ph: "" }]), cols, { rowH: 0.5 });
}

// ---------- 13 ----------
{
  const s = pres.addSlide();
  shell(s, 13, "PM·리더가 먼저 확보해야 하는 여섯 가지", "성실함으로 해결되지 않는 일은 프로젝트 구조로 확보합니다",
    "확인 시점 매트릭스: 6개 확보 항목 × 단계. 칸 표시는 원문 기준으로 기입. 교차 원칙: 기록된 가정과 변경 관리(PM 합의 시점).");
  const phases = ["착수", "설계", "구현", "검증", "배포"];
  const lw = 3.6, cw = (BW - lw) / phases.length, rh = 0.62;
  phases.forEach((p, j) => {
    const x = BX + lw + j * cw;
    box(s, x + 0.03, BY, cw - 0.06, 0.45, { fill: C.ink, r: 0.04 });
    T(s, p, { x, y: BY, w: cw, h: 0.45, fontSize: 11.5, bold: true, color: C.white, align: "center", valign: "middle" });
  });
  for (let i = 0; i < 6; i++) {
    const y = BY + 0.55 + i * (rh + 0.08);
    T(s, String(i + 1), { x: BX, y, w: 0.4, h: rh, fontSize: 18, bold: true, color: C.accent, valign: "middle" });
    ph(s, BX + 0.45, y, lw - 0.55, rh, "확보 항목", { valign: "middle" });
    phases.forEach((_, j) => {
      const x = BX + lw + j * cw;
      box(s, x + 0.03, y, cw - 0.06, rh, { fill: (i + j) % 2 ? C.band : "F7F9FB", r: 0.04 });
      s.addShape(pres.shapes.OVAL, { x: x + cw / 2 - 0.1, y: y + rh / 2 - 0.1, w: 0.2, h: 0.2, fill: { color: C.white }, line: { color: C.phLine, width: 0.75, dashType: "dash" } });
    });
  }
  T(s, "○ 확인 시점은 원문 기준으로 ● 표시", { x: BX + lw, y: BY + 0.55 + 6 * (rh + 0.08) - 0.02, w: BW - lw, h: 0.28, fontSize: 9, color: C.muted, align: "right" });
}

// ---------- 14 ----------
{
  const s = pres.addSlide();
  shell(s, 14, "출시 직전 점검과 다음 행동", "정상 흐름만 확인했다면 아직 끝난 것이 아닙니다",
    "8항목 점검 + 행동 원칙 4 + 다음 행동 3. v1 15·18·19장(출시 점검/함정 총정리/마무리) 통합. 교차 원칙: 실행 통제.");
  const lw = 5.6;
  T(s, "출시 직전 8개 점검 항목", { x: BX, y: BY, w: lw, h: 0.35, fontSize: 13, bold: true, color: C.ink });
  const iw = (lw - 0.2) / 2;
  for (let i = 0; i < 8; i++) {
    const col = Math.floor(i / 4), row = i % 4;
    const x = BX + col * (iw + 0.2), y = BY + 0.5 + row * 1.05;
    box(s, x, y + 0.08, 0.3, 0.3, { fill: C.white, line: C.ink, lw: 1.25, r: 0.04 });
    T(s, String(i + 1), { x: x + 0.4, y: y + 0.03, w: 0.4, h: 0.4, fontSize: 13, bold: true, color: C.accent, valign: "middle" });
    ph(s, x + 0.8, y, iw - 0.8, 0.9, "");
  }
  const rx = BX + lw + 0.35, rw = BW - lw - 0.35;
  box(s, rx, BY, rw, 2.55, { fill: C.band });
  T(s, "행동 원칙", { x: rx + 0.2, y: BY + 0.12, w: rw - 0.4, h: 0.35, fontSize: 13, bold: true, color: C.ink });
  for (let i = 0; i < 4; i++) {
    const x = rx + 0.2 + (i % 2) * ((rw - 0.5) / 2 + 0.1), y = BY + 0.55 + Math.floor(i / 2) * 0.97;
    ph(s, x, y, (rw - 0.5) / 2, 0.85, `원칙 ${i + 1}`);
  }
  const ny = BY + 2.75;
  box(s, rx, ny, rw, 1.95, { fill: C.ink });
  T(s, "다음 행동", { x: rx + 0.2, y: ny + 0.12, w: rw - 0.4, h: 0.35, fontSize: 13, bold: true, color: C.white });
  for (let i = 0; i < 3; i++) {
    const y = ny + 0.55 + i * 0.45;
    pill(s, rx + 0.2, y + 0.04, 0.32, 0.32, String(i + 1), { fill: C.accent, fs: 10 });
    T(s, "〔원문 입력〕", { x: rx + 0.65, y, w: rw - 0.85, h: 0.4, fontSize: 10, color: "CBD3DE", valign: "middle" });
  }
}

// ---------- 마무리 ----------
{
  const s = pres.addSlide();
  s.background = { color: C.ink };
  T(s, "마무리", { x: 0.9, y: 0.8, w: 6, h: 0.4, fontSize: 14, bold: true, color: C.accent });
  T(s, "단계마다 지금 무엇을 확인해야 하는지\n이 자료에서 찾아 실행하세요.", { x: 0.9, y: 1.3, w: 11.5, h: 1.5, fontSize: 30, bold: true, color: C.white });
  T(s, "역할별 읽기 경로", { x: 0.9, y: 3.35, w: 6, h: 0.4, fontSize: 14, bold: true, color: "CBD3DE" });
  const paths = [
    ["리더·임원 60초", [1, 2, 6, 8, 13, 14]],
    ["개발자", [4, 6, 7, 9, 10, 11]],
    ["PM·리더", [3, 4, 5, 12, 13, 14]],
  ];
  paths.forEach(([who, seq], i) => {
    const y = 3.95 + i * 0.95;
    T(s, who, { x: 0.9, y, w: 2.6, h: 0.6, fontSize: 15, bold: true, color: C.white, valign: "middle" });
    seq.forEach((n, j) => {
      const x = 3.6 + j * 1.25;
      s.addShape(pres.shapes.OVAL, { x, y, w: 0.6, h: 0.6, fill: { color: C.ink2 }, line: { color: C.accent, width: 1.5 } });
      T(s, String(n), { x, y, w: 0.6, h: 0.6, fontSize: 14, bold: true, color: C.white, align: "center", valign: "middle" });
      if (j < seq.length - 1) T(s, "→", { x: x + 0.62, y, w: 0.6, h: 0.6, fontSize: 14, color: C.side, align: "center", valign: "middle" });
    });
  });
  s.addNotes("마무리. 읽기 경로는 콘텐츠 전략 v2 §5.");
}

pres.writeFile({ fileName: path.join(__dirname, "first-ai-agent-team.pptx") }).then(f => console.log("wrote", f));
