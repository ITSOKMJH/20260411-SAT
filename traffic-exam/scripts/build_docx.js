// Usage: node scripts/build_docx.js data/YYYYMMDD.json output/
const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell,
  WidthType, ShadingType, AlignmentType, PageBreak, BorderStyle, Footer, PageNumber,
} = require('docx');

const [, , dataPath, outDir] = process.argv;
const d = JSON.parse(fs.readFileSync(dataPath, 'utf8'));
const FONT = '맑은 고딕';
const stars = (n) => '★'.repeat(n) + '☆'.repeat(3 - n);
const NAVY = '1E3A5F';
const W = 9026; // A4 width minus 2×1440 margins

const p = (text, opts = {}) => new Paragraph({
  spacing: { line: 360, after: opts.after ?? 80 },
  alignment: opts.align,
  children: [new TextRun({ text, font: FONT, size: opts.size ?? 22, bold: opts.bold, color: opts.color })],
});
const label = (k, v) => new Paragraph({
  spacing: { line: 300, after: 40 },
  indent: { left: 240 },
  children: [
    new TextRun({ text: k + '  ', font: FONT, size: 20, bold: true, color: NAVY }),
    new TextRun({ text: v, font: FONT, size: 20 }),
  ],
});
const cell = (text, w, header = false) => new TableCell({
  width: { size: w, type: WidthType.DXA },
  shading: header ? { type: ShadingType.CLEAR, color: 'auto', fill: NAVY } : undefined,
  margins: { top: 60, bottom: 60, left: 100, right: 100 },
  children: [new Paragraph({ children: [new TextRun({ text, font: FONT, size: 20, bold: header, color: header ? 'FFFFFF' : undefined })] })],
});
const table = (cols, rows) => new Table({
  width: { size: W, type: WidthType.DXA },
  columnWidths: cols,
  rows: [
    new TableRow({ tableHeader: true, children: rows[0].map((t, i) => cell(t, cols[i], true)) }),
    ...rows.slice(1).map((r) => new TableRow({ children: r.map((t, i) => cell(String(t), cols[i])) })),
  ],
});

const all = d.categories.flatMap((c) => c.qs.map((q, i) => ({ ...q, id: `${c.code}-${i + 1}`, cat: c.name })));
const count = (f) => all.filter(f).length;
const kids = [];

// Cover
kids.push(new Paragraph({ spacing: { before: 3000 }, children: [] }));
kids.push(p('교통기술사 일일 예상문제 50선', { size: 48, bold: true, align: AlignmentType.CENTER, color: NAVY, after: 300 }));
kids.push(p(d.dateKo, { size: 32, align: AlignmentType.CENTER, after: 200 }));
kids.push(p('“기출분석 × 정책동향 × 첨단모빌리티 × 미디어이슈”', { size: 24, align: AlignmentType.CENTER, after: 600 }));
kids.push(p(`목표 시험: ${d.exam}  |  D-${d.dday}`, { size: 22, align: AlignmentType.CENTER }));
kids.push(new Paragraph({ children: [new PageBreak()] }));

// Summary
kids.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: '출제 근거 분석 요약', font: FONT })] }));
kids.push(table([2600, 3226, 3200], [
  ['카테고리', '문항(난이도 ★/★★/★★★)', '유형 구성'],
  ...d.categories.map((c) => [
    `${c.code}. ${c.name}`,
    `5 (${[1, 2, 3].map((n) => c.qs.filter((q) => q.d === n).length).join('/')})`,
    [...new Set(c.qs.map((q) => q.t))].join(', '),
  ]),
  ['합계', `50 (${count((q) => q.d === 1)}/${count((q) => q.d === 2)}/${count((q) => q.d === 3)})`,
    `서술 ${count((q) => q.t === '서술형')} · 비교 ${count((q) => q.t === '비교·분석형')} · 계산 ${count((q) => q.t === '계산·수식형')} · 사례 ${count((q) => q.t === '사례적용형')}`],
]));
kids.push(p(''));
(d.analysis || []).forEach((a) => kids.push(label(a[0], a[1])));
kids.push(new Paragraph({ children: [new PageBreak()] }));

// Questions
d.categories.forEach((c) => {
  kids.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: `[${c.code}] ${c.name} (5문항)`, font: FONT })] }));
  c.qs.forEach((q, i) => {
    kids.push(new Paragraph({
      heading: HeadingLevel.HEADING_2,
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: NAVY, space: 2 } },
      children: [new TextRun({ text: `[${c.code}-${i + 1}] ${stars(q.d)}  ${q.t}`, font: FONT })],
    }));
    kids.push(p(q.q, { after: 100 }));
    kids.push(label('핵심 키워드', q.kw));
    kids.push(label('참고법규', q.law));
    kids.push(label('해외기준', q.intl));
    kids.push(label('출제 근거', q.basis));
  });
});

// Study guide
kids.push(new Paragraph({ children: [new PageBreak()] }));
kids.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: '오늘의 학습 가이드', font: FONT })] }));
const ids = (n) => all.filter((q) => q.d === n).map((q) => q.id).join(', ');
kids.push(label('필수 정복 (★☆☆)', ids(1) + ' → 개념 암기 우선'));
kids.push(label('집중 학습 (★★☆)', ids(2) + ' → 비교표 작성 연습'));
kids.push(label('심화 도전 (★★★)', ids(3) + ' → 창의적 제안 연습'));
kids.push(p('핵심 암기 포인트', { bold: true, color: NAVY }));
d.keyPoints.forEach((k, i) => kids.push(label(`${i + 1}.`, k)));
kids.push(label('내일 예고', d.tomorrow));

kids.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: '핵심 참고문헌·출처', font: FONT })] }));
(d.refs || []).forEach((r) => kids.push(label('·', r)));

const doc = new Document({
  styles: {
    default: { document: { run: { font: FONT, size: 22 } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { font: FONT, size: 32, bold: true, color: NAVY }, paragraph: { spacing: { before: 240, after: 160 }, keepNext: true, outlineLevel: 0 } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { font: FONT, size: 26, bold: true }, paragraph: { spacing: { before: 240, after: 100 }, keepNext: true, outlineLevel: 1 } },
    ],
  },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 18 })] })] }) },
    children: kids,
  }],
});

const out = path.join(outDir, `교통기술사_일일예상문제_${d.date.replace(/-/g, '')}.docx`);
Packer.toBuffer(doc).then((b) => { fs.writeFileSync(out, b); console.log(out); });
