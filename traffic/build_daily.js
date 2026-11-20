// 사용법: node build_daily.js data_YYYYMMDD.js  → 교통기술사_일일예상문제_YYYYMMDD.docx + .md
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow, TableCell,
  WidthType, ShadingType, BorderStyle, PageBreak, Footer, PageNumber,
} = require("docx");

const D = require(path.resolve(process.argv[2]));
const ymd = D.date.replace(/-/g, "");
const STAR = { 1: "★☆☆", 2: "★★☆", 3: "★★★" };
const LV = { 1: "기본", 2: "중급", 3: "심화" };
const FONT = "맑은 고딕";
const NAVY = "1E3A5F";

const p = (text, o = {}) => new Paragraph({ spacing: { line: 360, after: 80 }, ...o.para,
  children: [new TextRun({ text, font: FONT, size: o.size || 22, bold: o.bold, color: o.color })] });
const labeled = (label, text) => new Paragraph({ spacing: { line: 300, after: 40 }, indent: { left: 240 },
  children: [new TextRun({ text: label + "  ", font: FONT, size: 20, bold: true, color: NAVY }),
             new TextRun({ text, font: FONT, size: 20 })] });
const rule = () => new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: NAVY, space: 1 } }, spacing: { after: 120 } });

function table(header, rows, widths) {
  const total = widths.reduce((a, b) => a + b, 0);
  const border = { style: BorderStyle.SINGLE, size: 4, color: "999999" };
  const borders = { top: border, bottom: border, left: border, right: border };
  const cell = (t, i, head) => new TableCell({ width: { size: widths[i], type: WidthType.DXA }, borders,
    shading: head ? { type: ShadingType.CLEAR, fill: NAVY, color: "auto" } : undefined,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children: [new Paragraph({ children: [new TextRun({ text: String(t), font: FONT, size: 20, bold: head, color: head ? "FFFFFF" : undefined })] })] });
  return new Table({ width: { size: total, type: WidthType.DXA }, columnWidths: widths,
    rows: [new TableRow({ tableHeader: true, children: header.map((h, i) => cell(h, i, true)) }),
           ...rows.map(r => new TableRow({ children: r.map((c, i) => cell(c, i, false)) }))] });
}

const all = []; D.categories.forEach(c => c.qs.forEach((q, i) => all.push({ ...q, id: `${c.code}-${i + 1}`, cat: c.name })));
const typeCount = {}; all.forEach(q => typeCount[q.t] = (typeCount[q.t] || 0) + 1);
const byD = n => all.filter(q => q.d === n).map(q => q.id).join(", ");

const kids = [];
// 표지
kids.push(new Paragraph({ spacing: { before: 3000 } }));
kids.push(new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "교통기술사 일일 예상문제 50선", font: FONT, size: 48, bold: true, color: NAVY })] }));
kids.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 300 }, children: [new TextRun({ text: D.dateKo, font: FONT, size: 32 })] }));
kids.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 200 }, children: [new TextRun({ text: "기출분석 × 정책동향 × 첨단모빌리티 × 미디어이슈", font: FONT, size: 24, italics: true })] }));
kids.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 200 }, children: [new TextRun({ text: D.examNote, font: FONT, size: 20, color: "666666" })] }));
kids.push(new Paragraph({ children: [new PageBreak()] }));

// 출제 근거 분석
kids.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("출제 근거 분석 요약")] }));
kids.push(table(["구분", "반영 내용", "문항 수"], D.analysis, [1800, 6026, 1200]));
kids.push(p(""));
kids.push(table(["유형", "문항 수", "난이도", "문항 수"], [
  ["서술형", typeCount["서술형"] || 0, "★☆☆ 기본", all.filter(q => q.d === 1).length],
  ["비교·분석형", typeCount["비교·분석형"] || 0, "★★☆ 중급", all.filter(q => q.d === 2).length],
  ["계산·수식형", typeCount["계산·수식형"] || 0, "★★★ 심화", all.filter(q => q.d === 3).length],
  ["사례적용형", typeCount["사례적용형"] || 0, "합계", all.length],
], [2400, 2113, 2400, 2113]));
kids.push(new Paragraph({ children: [new PageBreak()] }));

// 문제
D.categories.forEach(c => {
  kids.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(`[카테고리 ${c.code}] ${c.name} (5문항)`)] }));
  c.qs.forEach((q, i) => {
    kids.push(new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(`[${c.code}-${i + 1}] ${STAR[q.d]} ${c.name} | ${q.t}`)] }));
    kids.push(p(q.q));
    kids.push(labeled("핵심 키워드", q.kw));
    kids.push(labeled("참고법규", q.law));
    kids.push(labeled("해외기준", q.intl));
    kids.push(labeled("출제 근거", q.basis));
    kids.push(rule());
  });
});

// 학습 가이드
kids.push(new Paragraph({ children: [new PageBreak()] }));
kids.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("오늘의 학습 가이드")] }));
kids.push(labeled("필수 정복 (★☆☆)", byD(1) + " → 개념 암기 우선"));
kids.push(labeled("집중 학습 (★★☆)", byD(2) + " → 비교표 작성 연습"));
kids.push(labeled("심화 도전 (★★★)", byD(3) + " → 창의적 제안 연습"));
kids.push(p("오늘의 핵심 암기 포인트", { bold: true }));
D.memo.forEach((m, i) => kids.push(labeled(`${i + 1}.`, m)));
kids.push(labeled("내일 예고 주제", D.tomorrow));
kids.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("핵심 참고문헌")] }));
D.refs.forEach((r, i) => kids.push(labeled(`[${i + 1}]`, r)));

const doc = new Document({
  styles: { default: { document: { run: { font: FONT, size: 22 } },
    heading1: { run: { font: FONT, size: 32, bold: true, color: NAVY }, paragraph: { spacing: { before: 240, after: 160 }, keepNext: true } },
    heading2: { run: { font: FONT, size: 26, bold: true }, paragraph: { spacing: { before: 200, after: 100 }, keepNext: true } } } },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 18 })] })] }) },
    children: kids }],
});

const out = path.join(__dirname, `교통기술사_일일예상문제_${ymd}.docx`);
Packer.toBuffer(doc).then(b => { fs.writeFileSync(out, b); console.log("wrote", out); });

// Markdown
let md = `# 📅 ${D.dateKo} 교통기술사 일일 예상문제 50선\n\n생성 기준: 기출분석 + 정책동향 + 첨단모빌리티 + 미디어이슈  \n${D.examNote}\n\n`;
D.categories.forEach(c => {
  md += `## [${c.code}] ${c.name}\n\n`;
  c.qs.forEach((q, i) => {
    md += `### [${c.code}-${i + 1}] ${STAR[q.d]} ${q.t}\n${q.q}\n\n- 📌 키워드: ${q.kw}\n- 📎 참고법규: ${q.law}\n- 🌐 해외기준: ${q.intl}\n- ⭐ 출제 근거: ${q.basis}\n\n`;
  });
});
md += `## 📊 출제 근거 분석\n\n| 구분 | 내용 | 문항 |\n|---|---|---|\n` + D.analysis.map(r => `| ${r.join(" | ")} |`).join("\n");
md += `\n\n## 📚 학습 전략\n\n- ✅ 필수 정복(★☆☆): ${byD(1)}\n- ⚡ 집중 학습(★★☆): ${byD(2)}\n- 🚀 심화 도전(★★★): ${byD(3)}\n\n**핵심 암기**\n` + D.memo.map((m, i) => `${i + 1}. ${m}`).join("\n");
md += `\n\n**내일 예고**: ${D.tomorrow}\n\n## 참고문헌\n` + D.refs.map(r => `- ${r}`).join("\n") + "\n";
fs.writeFileSync(path.join(__dirname, `교통기술사_일일예상문제_${ymd}.md`), md);
