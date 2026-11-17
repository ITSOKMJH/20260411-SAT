// usage: node build.js data-YYYYMMDD.js <outDir>
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell,
  WidthType, ShadingType, AlignmentType, PageBreak, BorderStyle,
} = require("docx");

const D = require(path.resolve(process.argv[2]));
const outDir = process.argv[3] || ".";
const ymd = D.date.replace(/-/g, "");
const star = (n) => "★".repeat(n) + "☆".repeat(3 - n);
const all = D.categories.flatMap((c) => c.qs.map((q) => ({ ...q, cat: c.name })));

// ---- sanity checks ----
const cnt = (f) => all.reduce((m, q) => ((m[f(q)] = (m[f(q)] || 0) + 1), m), {});
const lv = cnt((q) => q.lv), ty = cnt((q) => q.type);
console.log("total", all.length, "lv", lv, "type", ty);
if (all.length !== 50) throw new Error("need 50 questions");

// ---- Markdown ----
const md = [];
md.push(`# 📅 ${D.dateKo} 교통기술사 일일 예상문제 50선`, "");
md.push(`생성 기준: 기출분석 + 정책동향 + 첨단모빌리티 + 미디어이슈  `, `시험까지: ${D.exam}`, "");
for (const c of D.categories) {
  md.push(`## [${c.code}] ${c.name} (5문항)`, "");
  for (const q of c.qs) {
    md.push(`### [${q.id}] ${star(q.lv)} ${c.name} | ${q.type}`, "", q.q, "");
    md.push(`- 📌 핵심 키워드: ${q.kw}`, `- 📎 참고법규: ${q.law}`, `- 🌐 해외기준: ${q.intl}`, `- ⭐ 출제 근거: ${q.basis}`, "");
  }
}
md.push("## 📊 오늘의 출제 근거 분석", "", "| 구분 | 반영 내용 | 문항 |", "|---|---|---|");
D.analysis.forEach((r) => md.push(`| ${r[0]} | ${r[1]} | ${r[2]} |`));
const g = D.guide;
md.push("", "## 📚 오늘의 학습 전략", "", `- ✅ 필수 정복 (★☆☆): ${g.must}`, `- ⚡ 집중 학습 (★★☆): ${g.focus}`, `- 🚀 심화 도전 (★★★): ${g.deep}`, "", "**💡 핵심 암기 포인트**", "");
g.points.forEach((p, i) => md.push(`${i + 1}. ${p}`));
md.push("", `**📅 내일 예고:** ${g.tomorrow}`, "", "## 핵심 참고문헌", "");
D.refs.forEach((r) => md.push(`- ${r}`));
md.push("", "> 계산형 문항의 수치는 가상 수치이며, 법규·지침 명칭은 최신 개정본 확인이 필요합니다.");
fs.writeFileSync(path.join(outDir, `교통기술사_일일예상문제_${ymd}.md`), md.join("\n") + "\n");

// ---- DOCX ----
const F = "맑은 고딕";
const p = (text, o = {}) => new Paragraph({ spacing: { line: 360, after: 80 }, ...o, children: [new TextRun({ text, font: F, size: 22, ...(o.run || {}) })] });
const label = (k, v) => new Paragraph({ spacing: { line: 300, after: 40 }, indent: { left: 280 }, children: [new TextRun({ text: k + " ", font: F, size: 20, bold: true, color: "1E3A5F" }), new TextRun({ text: v, font: F, size: 20 })] });
const W = 9026;
const cell = (t, w, head) => new TableCell({
  width: { size: w, type: WidthType.DXA },
  shading: head ? { type: ShadingType.CLEAR, color: "auto", fill: "1E3A5F" } : undefined,
  margins: { top: 60, bottom: 60, left: 100, right: 100 },
  children: [new Paragraph({ children: [new TextRun({ text: t, font: F, size: 20, bold: head, color: head ? "FFFFFF" : undefined })] })],
});
const table = (cols, rows) => new Table({
  width: { size: W, type: WidthType.DXA }, columnWidths: cols,
  rows: rows.map((r, i) => new TableRow({ tableHeader: i === 0, children: r.map((t, j) => cell(String(t), cols[j], i === 0)) })),
});

const kids = [];
kids.push(new Paragraph({ spacing: { before: 3000 } }));
kids.push(new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "교통기술사 일일 예상문제 50선", font: F, size: 48, bold: true, color: "1E3A5F" })] }));
kids.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 300 }, children: [new TextRun({ text: D.dateKo, font: F, size: 32 })] }));
kids.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 300 }, children: [new TextRun({ text: "기출분석 × 정책동향 × 첨단모빌리티", font: F, size: 26, italics: true })] }));
kids.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 300 }, children: [new TextRun({ text: D.exam, font: F, size: 20, color: "666666" })] }));
kids.push(new Paragraph({ children: [new PageBreak()] }));

kids.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("출제 근거 분석 요약")] }));
kids.push(table([1800, 6226, 1000], [["구분", "반영 내용", "문항"], ...D.analysis]));
kids.push(p(""));
kids.push(table([3000, 3013, 3013], [["구분", "난이도 분포", "유형 분포"],
  ["문항 수", `★☆☆ ${lv[1]} / ★★☆ ${lv[2]} / ★★★ ${lv[3]}`, Object.entries(ty).map(([k, v]) => `${k} ${v}`).join(", ")]]));

for (const c of D.categories) {
  kids.push(new Paragraph({ children: [new PageBreak()] }));
  kids.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(`[${c.code}] ${c.name} (5문항)`)] }));
  for (const q of c.qs) {
    kids.push(new Paragraph({ heading: HeadingLevel.HEADING_2, border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "1E3A5F", space: 1 } }, children: [new TextRun(`[${q.id}] ${star(q.lv)}  ${q.type}`)] }));
    kids.push(p(q.q));
    kids.push(label("📌 핵심 키워드:", q.kw));
    kids.push(label("📎 참고법규:", q.law));
    kids.push(label("🌐 해외기준:", q.intl));
    kids.push(label("⭐ 출제 근거:", q.basis));
  }
}

kids.push(new Paragraph({ children: [new PageBreak()] }));
kids.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("오늘의 학습 가이드")] }));
kids.push(table([2200, 6826], [["단계", "문항"], ["✅ 필수 정복 ★☆☆", g.must], ["⚡ 집중 학습 ★★☆", g.focus], ["🚀 심화 도전 ★★★", g.deep]]));
kids.push(new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("💡 핵심 암기 포인트")] }));
g.points.forEach((t, i) => kids.push(p(`${i + 1}. ${t}`)));
kids.push(new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("📅 내일 예고")] }));
kids.push(p(g.tomorrow));
kids.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("핵심 참고문헌")] }));
D.refs.forEach((r) => kids.push(p("· " + r)));
kids.push(p("※ 계산형 문항 수치는 가상 수치이며, 법규·지침은 최신 개정본을 확인할 것.", { run: { color: "888888", size: 18 } }));

const doc = new Document({
  styles: {
    default: { document: { run: { font: F, size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: F, size: 32, bold: true, color: "1E3A5F" }, paragraph: { spacing: { before: 240, after: 160 }, keepNext: true, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: F, size: 26, bold: true }, paragraph: { spacing: { before: 280, after: 100 }, keepNext: true, outlineLevel: 1 } },
    ],
  },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } }, children: kids }],
});
Packer.toBuffer(doc).then((b) => {
  const f = path.join(outDir, `교통기술사_일일예상문제_${ymd}.docx`);
  fs.writeFileSync(f, b);
  console.log("wrote", f);
});
