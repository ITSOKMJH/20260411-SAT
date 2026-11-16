// Usage: node build.js <questions.json> <outDir>
// Generates 교통기술사_일일예상문제_YYYYMMDD.docx and .md from the question JSON.
const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, WidthType, ShadingType, BorderStyle, PageBreak,
} = require('docx');

const [, , jsonPath, outDir] = process.argv;
const data = JSON.parse(fs.readFileSync(jsonPath, 'utf8'));
const ymd = data.date.replace(/-/g, '');
const FONT = '맑은 고딕';
const stars = (n) => '★'.repeat(n) + '☆'.repeat(3 - n);
const all = data.categories.flatMap((c) => c.qs.map((q, i) => ({ ...q, no: `${c.id}-${i + 1}`, cat: c.name })));

// --- sanity checks against the skill's distribution rules
const count = (f) => all.filter(f).length;
const stats = {
  total: all.length,
  서술형: count((q) => q.type === '서술형'), '비교·분석형': count((q) => q.type === '비교·분석형'),
  '계산·수식형': count((q) => q.type === '계산·수식형'), 사례적용형: count((q) => q.type === '사례적용형'),
  lv1: count((q) => q.lv === 1), lv2: count((q) => q.lv === 2), lv3: count((q) => q.lv === 3),
};
console.log(stats);

const p = (text, opts = {}) => new Paragraph({ children: [new TextRun({ text, font: FONT, ...opts.run })], ...opts.para });
const border = { style: BorderStyle.SINGLE, size: 4, color: '9AA5B1' };
const borders = { top: border, bottom: border, left: border, right: border };
function table(headers, rows, widths) {
  const total = widths.reduce((a, b) => a + b, 0);
  const cell = (t, w, head) => new TableCell({
    width: { size: w, type: WidthType.DXA }, borders,
    shading: head ? { type: ShadingType.CLEAR, color: 'auto', fill: '1E3A5F' } : undefined,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children: [new Paragraph({ children: [new TextRun({ text: String(t), font: FONT, size: 19, bold: head, color: head ? 'FFFFFF' : undefined })] })],
  });
  return new Table({
    width: { size: total, type: WidthType.DXA }, columnWidths: widths,
    rows: [new TableRow({ tableHeader: true, children: headers.map((h, i) => cell(h, widths[i], true)) }),
      ...rows.map((r) => new TableRow({ children: r.map((t, i) => cell(t, widths[i], false)) }))],
  });
}

const children = [];
// Cover
children.push(p('', { para: { spacing: { before: 3000 } } }));
children.push(p('교통기술사 일일 예상문제 50선', { run: { size: 48, bold: true, color: '1E3A5F' }, para: { alignment: AlignmentType.CENTER } }));
children.push(p(data.dateKo, { run: { size: 30 }, para: { alignment: AlignmentType.CENTER, spacing: { before: 400 } } }));
children.push(p('“기출분석 × 정책동향 × 첨단모빌리티 × 미디어이슈”', { run: { size: 24, italics: true }, para: { alignment: AlignmentType.CENTER, spacing: { before: 400 } } }));
children.push(p(data.dday, { run: { size: 22, color: 'B03A2E' }, para: { alignment: AlignmentType.CENTER, spacing: { before: 400 } } }));
children.push(new Paragraph({ children: [new PageBreak()] }));

// Analysis summary
children.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: '출제 근거 분석 요약', font: FONT })] }));
children.push(table(['구분', '반영 내용', '문항'], data.analysis.map((a) => [a.k, a.v, a.n]), [1800, 6026, 1200]));
children.push(p('', {}));
children.push(table(['유형/난이도', '문항 수'], [
  ['서술형', stats['서술형']], ['비교·분석형', stats['비교·분석형']], ['계산·수식형', stats['계산·수식형']], ['사례적용형', stats['사례적용형']],
  ['★☆☆ 기본', stats.lv1], ['★★☆ 중급', stats.lv2], ['★★★ 심화', stats.lv3],
], [5000, 4026]));
children.push(new Paragraph({ children: [new PageBreak()] }));

// Questions
for (const c of data.categories) {
  children.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: `[카테고리 ${c.id}] ${c.name} (5문항)`, font: FONT })] }));
  c.qs.forEach((q, i) => {
    children.push(new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun({ text: `[${c.id}-${i + 1}] ${stars(q.lv)}  ${q.type}`, font: FONT })] }));
    children.push(p(q.text));
    for (const [k, v] of [['핵심 키워드', q.kw], ['참고법규', q.law], ['해외기준', q.intl], ['출제 근거', q.basis]]) {
      children.push(new Paragraph({ indent: { left: 360 }, spacing: { line: 300 }, children: [
        new TextRun({ text: `${k}: `, font: FONT, bold: true, size: 19, color: '1E3A5F' }),
        new TextRun({ text: v, font: FONT, size: 19 })] }));
    }
  });
}
children.push(new Paragraph({ children: [new PageBreak()] }));

// Study guide
const g = data.guide;
children.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: '오늘의 학습 가이드', font: FONT })] }));
children.push(p(`필수 정복(★☆☆): ${g.basic}`));
children.push(p(`집중 학습(★★☆): ${g.mid}`));
children.push(p(`심화 도전(★★★): ${g.adv}`));
children.push(new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun({ text: '핵심 암기 포인트', font: FONT })] }));
g.points.forEach((t, i) => children.push(p(`${i + 1}. ${t}`)));
children.push(new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun({ text: '내일 예고', font: FONT })] }));
children.push(p(g.tomorrow));
children.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: '핵심 참고문헌·출처', font: FONT })] }));
data.refs.forEach((r) => children.push(p(r, { run: { size: 19 } })));

const doc = new Document({
  styles: {
    default: { document: { run: { font: FONT, size: 22 }, paragraph: { spacing: { line: 360 } } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { font: FONT, size: 32, bold: true, color: '1E3A5F' }, paragraph: { spacing: { before: 360, after: 160 }, keepNext: true, outlineLevel: 0 } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { font: FONT, size: 26, bold: true }, paragraph: { spacing: { before: 240, after: 80 }, keepNext: true, outlineLevel: 1 } },
    ],
  },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } }, children }],
});

fs.mkdirSync(outDir, { recursive: true });
Packer.toBuffer(doc).then((buf) => {
  const f = path.join(outDir, `교통기술사_일일예상문제_${ymd}.docx`);
  fs.writeFileSync(f, buf);
  console.log('wrote', f);
});

// Markdown version
const L = [`# 📅 ${data.dateKo} 교통기술사 일일 예상문제 50선`, '', `생성 기준: 기출분석 + 정책동향 + 첨단모빌리티 + 미디어이슈  `, `시험까지: ${data.dday}`, ''];
for (const c of data.categories) {
  L.push(`## [${c.id}] ${c.name}`, '');
  c.qs.forEach((q, i) => {
    L.push(`### [${c.id}-${i + 1}] ${stars(q.lv)} ${c.name} | ${q.type}`, '', q.text, '',
      `- 📌 핵심 키워드: ${q.kw}`, `- 📎 참고법규: ${q.law}`, `- 🌐 해외기준: ${q.intl}`, `- ⭐ 출제 근거: ${q.basis}`, '');
  });
}
L.push('## 📊 오늘의 출제 근거 분석', '');
data.analysis.forEach((a) => L.push(`- 【${a.k}】 ${a.v} → ${a.n}`));
L.push('', '## 📚 오늘의 학습 전략', '', `- ✅ 필수 정복(★☆☆): ${g.basic}`, `- ⚡ 집중 학습(★★☆): ${g.mid}`, `- 🚀 심화 도전(★★★): ${g.adv}`, '', '**💡 핵심 암기 포인트**', '');
g.points.forEach((t, i) => L.push(`${i + 1}. ${t}`));
L.push('', `**📅 내일 예고:** ${g.tomorrow}`, '', '## 참고 출처', '');
data.refs.forEach((r) => L.push(`- ${r}`));
fs.writeFileSync(path.join(outDir, `교통기술사_일일예상문제_${ymd}.md`), L.join('\n') + '\n');
