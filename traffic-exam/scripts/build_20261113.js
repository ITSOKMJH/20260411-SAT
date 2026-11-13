const fs = require("fs");
const path = require("path");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ShadingType,
  AlignmentType, HeadingLevel, PageBreak, BorderStyle } = require("docx");
const { CATS, Q } = require("./q_20261113.js");

const DATE = "2026년 11월 13일";
const OUT = path.join(__dirname, "..", "2026-11-13");
fs.mkdirSync(OUT, { recursive: true });

// ---- sanity checks ----
const cnt = (f) => Q.reduce((m, q) => ((m[f(q)] = (m[f(q)] || 0) + 1), m), {});
const types = cnt((q) => q.t), diffs = cnt((q) => q.d), cats = cnt((q) => q.id[0]);
console.log(Q.length, types, diffs, cats);
if (Q.length !== 50) throw new Error("need 50");

const ANALYSIS = [
  ["기출 반영", "최근 10년 고빈도 주제(포화교통류율, 중력모형, 설계속도, 준공영제, C-ITS, SAE 레벨, 교통정온화, VOT, 교통영향평가, 디지털트윈 등)", "20"],
  ["정책 반영", "대광위 신도시 도로사업 16곳 점검(2026.11), 무인 자율주행차 가이드라인(2026.7), 자율주행법 개정(2026.12.23 시행), 2030 모빌리티 혁신성장 로드맵, 국토부 2027 예산안 69조, K-Safer", "11"],
  ["첨단기술", "레벨4 자율주행·K-UAM 시범운용 모델·C-ITS 통신방식·인프라 협력 자율주행", "7"],
  ["미디어 이슈", "강릉 ITS 세계총회(2026.10), 광주 자율차 200대 실증, 고령운전자 사고, PM 규제, C-ITS 안전망 국회 토론회", "7"],
  ["신규 트렌드·해외", "TRB 2026 DRL 신호제어·사이버보안, 파운데이션 모델 노변감지, 물류자동화, 주차로봇·스마트시티", "5"],
];

const GUIDE = {
  must: "A-1, A-2, B-1, B-2, C-1, C-2, D-1, D-2, E-1, F-1, G-1, G-2, H-1, I-1, J-1 → 개념·정의·수식 암기 우선",
  focus: "A-3, A-4, B-3, B-4, C-3, C-4, D-3, D-4, E-2~E-4, F-2~F-4, G-3, G-4, H-2~H-4, I-2~I-4, J-2~J-4 → 비교표·절차도 작성 연습",
  deep: "A-5, B-5, C-5, D-5, E-5, F-5, G-5, H-5, I-5, J-5 → 최신 정책·사례 인용 + 창의적 제안",
  memo: [
    "Webster 최적주기 C₀ = (1.5L + 5) / (1 − Y), L: 총손실시간, Y: 임계 v/s 합",
    "제동거리 d = V² / (2g(f ± G)) — 내리막은 (f − G), 정지시거 = 반응거리(V·t) + 제동거리",
    "자율주행법 개정(2026.12.23 시행): 시·도지사도 시범운행지구 지정 가능 → 지방정부 주도 실증 확대",
  ],
  tomorrow: "B(교통수요·계획: 활동기반모형·통행시간 신뢰성) / H(교통투자·평가: 민자사업 VfM·실물옵션) 집중",
};

const REFS = [
  "국토교통부, 도로용량편람(2023)",
  "국토교통부, 교통시설 투자평가지침(제7판)",
  "국토교통부, 무인 자율주행차 안전운행 요건 가이드라인(2026.7)",
  "국토교통부, 2030 모빌리티 혁신성장 로드맵(2026)",
  "대도시권광역교통위원회, 신도시 교통대책 조기이행 점검 보도자료(2026.11) — eiec.kdi.re.kr/policy/materialView.do?num=289316",
  "2026 강릉 ITS 세계총회(2026.10.19~23) 학술세션 자료 — news.mtn.co.kr/news-detail/2026091922550017631",
  "Transportation Research Board, HCM 7th Edition (2022)",
  "TRB 105th Annual Meeting (2026.1) 발표논문 — TRID: trid.trb.org/View/2761049 (Game Theory+AI 신호제어 리뷰)",
  "AASHTO, Highway Safety Manual / Green Book 8th Ed.",
  "SAE J3016 (2021) Taxonomy of Driving Automation",
];

// ---- DOCX helpers ----
const F = "맑은 고딕";
const run = (text, o = {}) => new TextRun({ text, font: F, size: 22, ...o });
const p = (children, o = {}) => new Paragraph({ children: Array.isArray(children) ? children : [run(children)], spacing: { line: 360, after: 80 }, ...o });
const H1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: t, font: F, bold: true, size: 32 })] });
const H2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun({ text: t, font: F, bold: true, size: 26 })] });
const border = { style: BorderStyle.SINGLE, size: 4, color: "999999" };
const borders = { top: border, bottom: border, left: border, right: border };
function table(widths, header, rows) {
  const total = widths.reduce((a, b) => a + b, 0);
  const cell = (t, i, hdr) => new TableCell({
    borders, width: { size: widths[i], type: WidthType.DXA },
    shading: hdr ? { fill: "1E3A5F", type: ShadingType.CLEAR, color: "auto" } : undefined,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children: [new Paragraph({ children: [run(t, hdr ? { bold: true, color: "FFFFFF", size: 20 } : { size: 20 })] })],
  });
  return new Table({
    width: { size: total, type: WidthType.DXA }, columnWidths: widths,
    rows: [new TableRow({ tableHeader: true, children: header.map((h, i) => cell(h, i, true)) }),
      ...rows.map((r) => new TableRow({ children: r.map((c, i) => cell(c, i, false)) }))],
  });
}
const rule = () => new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "1E3A5F", space: 1 } }, children: [] });

const children = [];
// cover
children.push(new Paragraph({ spacing: { before: 3000 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: "교통기술사 일일 예상문제 50선", font: F, bold: true, size: 48, color: "1E3A5F" })] }));
children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 400 }, children: [run(DATE, { size: 32 })] }));
children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 400 }, children: [run("“기출분석 × 정책동향 × 첨단모빌리티 × 미디어이슈”", { size: 26, italics: true })] }));
children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 800 }, children: [run("시험까지 D-78 (2027년 제1회 필기 추정일 2027-01-30 기준, Q-Net 공고 확인 필요)", { size: 20, color: "666666" })] }));
children.push(new Paragraph({ children: [new PageBreak()] }));

// analysis
children.push(H1("출제 근거 분석 요약"));
children.push(table([1800, 6026, 1200], ["구분", "반영 내용", "문항 수"], ANALYSIS));
children.push(p(""));
children.push(table([3008, 3009, 3009], ["유형", "난이도", "카테고리"], [[
  Object.entries(types).map(([k, v]) => `${k} ${v}`).join(" / "),
  Object.entries(diffs).map(([k, v]) => `${k} ${v}`).join(" / "),
  "A~J 각 5문항 (총 50)",
]]));
children.push(p([run("※ 분류는 중복 집계 가능(한 문항이 정책·미디어 동시 반영). 계산형 수치는 모두 가상 수치임.", { size: 18, color: "666666" })]));

// questions
for (const c of Object.keys(CATS)) {
  children.push(new Paragraph({ children: [new PageBreak()] }));
  children.push(H1(`[카테고리 ${c}] ${CATS[c]} (5문항)`));
  for (const q of Q.filter((x) => x.id[0] === c)) {
    children.push(H2(`[${q.id}] ${q.d} ${CATS[c]} | ${q.t}`));
    children.push(p(q.q));
    children.push(p([run("📌 핵심 키워드: ", { bold: true }), run(q.kw)]));
    children.push(p([run("📎 참고법규: ", { bold: true }), run(q.law)]));
    children.push(p([run("🌐 해외기준: ", { bold: true }), run(q.intl)]));
    children.push(p([run("⭐ 출제 근거: ", { bold: true }), run(q.basis)]));
    children.push(rule());
  }
}

// guide
children.push(new Paragraph({ children: [new PageBreak()] }));
children.push(H1("오늘의 학습 가이드"));
children.push(table([2200, 6826], ["구분", "대상 문항 및 전략"], [
  ["✅ 필수 정복 ★☆☆", GUIDE.must], ["⚡ 집중 학습 ★★☆", GUIDE.focus], ["🚀 심화 도전 ★★★", GUIDE.deep]]));
children.push(H2("💡 오늘의 핵심 암기 포인트"));
GUIDE.memo.forEach((m, i) => children.push(p(`${i + 1}. ${m}`)));
children.push(H2("📅 내일 예고 주제"));
children.push(p(GUIDE.tomorrow));
children.push(H1("핵심 참고문헌"));
REFS.forEach((r, i) => children.push(p(`[${i + 1}] ${r}`)));

const doc = new Document({
  styles: {
    default: { document: { run: { font: F, size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: F, size: 32, bold: true, color: "1E3A5F" }, paragraph: { spacing: { before: 240, after: 160 }, keepNext: true, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: F, size: 26, bold: true }, paragraph: { spacing: { before: 240, after: 100 }, keepNext: true, outlineLevel: 1 } },
    ],
  },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } }, children }],
});
Packer.toBuffer(doc).then((b) => {
  const f = path.join(OUT, "교통기술사_일일예상문제_20261113.docx");
  fs.writeFileSync(f, b);
  console.log("wrote", f);
});

// ---- Markdown ----
const md = [];
md.push(`# 📅 ${DATE} 교통기술사 일일 예상문제 50선`, "", "생성 기준: 기출분석 + 정책동향 + 첨단모빌리티 + 미디어이슈  ", "시험까지: D-78 (2027년 제1회 필기 추정일 2027-01-30 기준, Q-Net 공고 확인 필요)", "");
for (const c of Object.keys(CATS)) {
  md.push(`## [${c}] ${CATS[c]}`, "");
  for (const q of Q.filter((x) => x.id[0] === c)) {
    md.push(`### [${q.id}] ${q.d} ${CATS[c]} | ${q.t}`, "", q.q, "",
      `- 📌 핵심 키워드: ${q.kw}`, `- 📎 참고법규: ${q.law}`, `- 🌐 해외기준: ${q.intl}`, `- ⭐ 출제 근거: ${q.basis}`, "");
  }
}
md.push("## 📊 오늘의 출제 근거 분석", "", "| 구분 | 반영 내용 | 문항 수 |", "|---|---|---|", ...ANALYSIS.map((r) => `| ${r.join(" | ")} |`), "",
  `유형: ${Object.entries(types).map(([k, v]) => `${k} ${v}`).join(" / ")} · 난이도: ${Object.entries(diffs).map(([k, v]) => `${k} ${v}`).join(" / ")}`, "",
  "## 📚 오늘의 학습 전략", "", `- ✅ 필수 정복 (★☆☆): ${GUIDE.must}`, `- ⚡ 집중 학습 (★★☆): ${GUIDE.focus}`, `- 🚀 심화 도전 (★★★): ${GUIDE.deep}`, "",
  "**💡 핵심 암기 포인트**", "", ...GUIDE.memo.map((m, i) => `${i + 1}. ${m}`), "", `**📅 내일 예고:** ${GUIDE.tomorrow}`, "",
  "## 핵심 참고문헌", "", ...REFS.map((r, i) => `${i + 1}. ${r}`), "");
fs.writeFileSync(path.join(OUT, "교통기술사_일일예상문제_20261113.md"), md.join("\n"));
