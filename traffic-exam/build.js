// Usage: node build.js data/YYYYMMDD.js  → out/교통기술사_일일예상문제_YYYYMMDD.docx + .md
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType,
  ShadingType, HeadingLevel, AlignmentType, PageBreak, BorderStyle,
} = require("docx");

const D = require(path.resolve(process.argv[2]));
const stamp = D.date.replace(/-/g, "");
const FONT = "맑은 고딕";
const star = (d) => "★".repeat(d) + "☆".repeat(3 - d);
const all = D.categories.flatMap((c) => c.qs.map((q, i) => ({ ...q, id: `${c.code}-${i + 1}`, cat: c.name })));

// sanity checks against the skill's distribution rules
const count = (k) => all.reduce((m, q) => ((m[q[k]] = (m[q[k]] || 0) + 1), m), {});
const types = count("t"), diffs = count("d");
if (all.length !== 50) throw new Error("need 50 questions, got " + all.length);
console.log("types", types, "diffs", diffs);

const p = (text, opts = {}) => new Paragraph({ spacing: { line: 360 }, ...opts, children: [new TextRun({ text, font: FONT, ...(opts.run || {}) })] });
const label = (k, v) => new Paragraph({ spacing: { line: 300 }, children: [new TextRun({ text: k + " ", bold: true, font: FONT, size: 20 }), new TextRun({ text: v, font: FONT, size: 20 })] });

const W = 9026;
function table(header, rows, widths) {
  const cell = (t, i, head) => new TableCell({
    width: { size: widths[i], type: WidthType.DXA },
    shading: head ? { type: ShadingType.CLEAR, fill: "1E3A5F", color: "auto" } : undefined,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children: [new Paragraph({ children: [new TextRun({ text: t, font: FONT, size: 20, bold: head, color: head ? "FFFFFF" : undefined })] })],
  });
  return new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: widths,
    rows: [new TableRow({ tableHeader: true, children: header.map((h, i) => cell(h, i, true)) }),
      ...rows.map((r) => new TableRow({ children: r.map((t, i) => cell(String(t), i, false)) }))],
  });
}

const kids = [];
// cover
kids.push(new Paragraph({ spacing: { before: 3000 } }));
kids.push(p("교통기술사 일일 예상문제 50선", { alignment: AlignmentType.CENTER, run: { size: 48, bold: true, color: "1E3A5F" } }));
kids.push(p(D.dateKo, { alignment: AlignmentType.CENTER, run: { size: 32 } }));
kids.push(p("기출분석 × 정책동향 × 첨단모빌리티 × 미디어이슈", { alignment: AlignmentType.CENTER, run: { size: 24, italics: true } }));
kids.push(p(`시험까지: ${D.dday}`, { alignment: AlignmentType.CENTER, run: { size: 20 } }));
kids.push(p(D.rotation, { alignment: AlignmentType.CENTER, run: { size: 20 } }));
kids.push(new Paragraph({ children: [new PageBreak()] }));

// analysis
kids.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("오늘의 출제 근거 분석")] }));
kids.push(table(["구분", "반영 내용", "문항"], D.analysis, [1800, 6326, 900]));
kids.push(p(""));
kids.push(table(["유형", "문항 수"], Object.entries(types), [6026, 3000]));
kids.push(p(""));
kids.push(table(["난이도", "문항 수"], Object.entries(diffs).map(([d, n]) => [star(+d), n]), [6026, 3000]));

// questions
for (const c of D.categories) {
  kids.push(new Paragraph({ children: [new PageBreak()] }));
  kids.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(`[카테고리 ${c.code}] ${c.name} (5문항)`)] }));
  c.qs.forEach((q, i) => {
    kids.push(new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(`[${c.code}-${i + 1}] ${star(q.d)}  ${q.t}`)] }));
    kids.push(p(q.q));
    kids.push(label("핵심 키워드:", q.kw));
    kids.push(label("참고법규:", q.law));
    kids.push(label("해외기준:", q.intl));
    kids.push(label("출제 근거:", q.basis));
    kids.push(new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "999999", space: 4 } } }));
  });
}

// guide + refs
kids.push(new Paragraph({ children: [new PageBreak()] }));
kids.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("오늘의 학습 가이드")] }));
kids.push(label("필수 정복(★☆☆):", D.guide.must + " → 개념 암기 우선"));
kids.push(label("집중 학습(★★☆):", D.guide.focus + " → 비교표 작성 연습"));
kids.push(label("심화 도전(★★★):", D.guide.adv + " → 창의적 제안 연습"));
kids.push(p("핵심 암기 포인트", { run: { bold: true } }));
D.guide.memo.forEach((m, i) => kids.push(p(`${i + 1}. ${m}`)));
kids.push(label("내일 예고:", D.guide.tomorrow));
kids.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("핵심 참고문헌")] }));
D.refs.forEach((r, i) => kids.push(p(`[${i + 1}] ${r}`, { run: { size: 20 } })));
kids.push(p("※ '가상 수치 적용' 표기 문항의 수치는 학습용 가정값입니다. 법규·지침은 최신 개정본을 다시 확인하세요.", { run: { size: 18, italics: true } }));

const doc = new Document({
  styles: {
    default: { document: { run: { font: FONT, size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: FONT, size: 32, bold: true, color: "1E3A5F" }, paragraph: { spacing: { before: 240, after: 160 }, keepNext: true, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: FONT, size: 26, bold: true }, paragraph: { spacing: { before: 200, after: 80 }, keepNext: true, outlineLevel: 1 } },
    ],
  },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } }, children: kids }],
});

const out = path.join(__dirname, "out");
fs.mkdirSync(out, { recursive: true });
Packer.toBuffer(doc).then((buf) => {
  const f = path.join(out, `교통기술사_일일예상문제_${stamp}.docx`);
  fs.writeFileSync(f, buf);
  // markdown mirror for quick reading on phone/GitHub
  let md = `# 📅 ${D.dateKo} 교통기술사 일일 예상문제 50선\n\n생성 기준: 기출분석 + 정책동향 + 첨단모빌리티 + 미디어이슈  \n시험까지: ${D.dday}\n\n`;
  for (const c of D.categories) {
    md += `## [${c.code}] ${c.name}\n\n`;
    c.qs.forEach((q, i) => {
      md += `**[${c.code}-${i + 1}] ${star(q.d)} ${q.t}**  \n${q.q}\n\n- 📌 ${q.kw}\n- 📎 ${q.law}\n- 🌐 ${q.intl}\n- ⭐ ${q.basis}\n\n`;
    });
  }
  md += `## 📊 출제 근거 분석\n\n| 구분 | 내용 | 문항 |\n|---|---|---|\n` + D.analysis.map((r) => `| ${r.join(" | ")} |`).join("\n");
  md += `\n\n## 📚 학습 가이드\n\n- ✅ 필수: ${D.guide.must}\n- ⚡ 집중: ${D.guide.focus}\n- 🚀 심화: ${D.guide.adv}\n\n` + D.guide.memo.map((m, i) => `${i + 1}. ${m}`).join("\n") + `\n\n📅 내일 예고: ${D.guide.tomorrow}\n\n## 참고문헌\n\n` + D.refs.map((r) => `- ${r}`).join("\n") + "\n";
  fs.writeFileSync(path.join(out, `교통기술사_일일예상문제_${stamp}.md`), md);
  console.log("wrote", f);
});
