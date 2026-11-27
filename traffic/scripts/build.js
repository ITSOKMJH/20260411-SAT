// Usage: node build.js <data.json> <out_dir>
const fs = require('fs');
const path = require('path');
const { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell,
  WidthType, ShadingType, AlignmentType, PageBreak, BorderStyle } = require('/opt/node22/lib/node_modules/docx');

const d = JSON.parse(fs.readFileSync(process.argv[2], 'utf8'));
const outDir = process.argv[3];
const CATS = { A:'교통신호·교차로', B:'교통수요·계획', C:'도로설계·용량', D:'대중교통·환승', E:'ITS·스마트교통',
  F:'자율주행·미래모빌리티', G:'교통안전·도로안전', H:'교통투자·평가', I:'교통정책·법규', J:'AI·빅데이터·융합' };
const F = '맑은 고딕';
const P = (text, o = {}) => new Paragraph({ spacing: { line: 360, after: 80 }, alignment: o.align,
  children: [new TextRun({ text, font: F, size: o.size || 22, bold: o.bold, color: o.color })] });
const cell = (t, hdr, w) => new TableCell({ width: { size: w, type: WidthType.DXA },
  shading: hdr ? { type: ShadingType.CLEAR, fill: '1E3A5F', color: 'auto' } : undefined,
  children: [new Paragraph({ children: [new TextRun({ text: t, font: F, size: 20, bold: hdr, color: hdr ? 'FFFFFF' : undefined })] })] });
const table = (rows, widths) => new Table({ width: { size: widths.reduce((a, b) => a + b), type: WidthType.DXA }, columnWidths: widths,
  rows: rows.map((r, i) => new TableRow({ children: r.map((t, j) => cell(t, i === 0, widths[j])) })) });

const kids = [];
kids.push(P('', {}), P('', {}), P('교통기술사 일일 예상문제 50선', { size: 48, bold: true, align: AlignmentType.CENTER, color: '1E3A5F' }),
  P(d.dateKo, { size: 32, align: AlignmentType.CENTER }),
  P('"기출분석 × 정책동향 × 첨단모빌리티"', { size: 26, align: AlignmentType.CENTER }),
  P(d.dday, { size: 22, align: AlignmentType.CENTER }),
  new Paragraph({ children: [new PageBreak()] }));
kids.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: '출제 근거 분석 요약', font: F, size: 32, bold: true })] }));
kids.push(table([['구분', '반영 내용', '문항'], ...d.analysis], [1800, 6226, 1000]));
const lv = s => d.questions.filter(q => q.diff === s).length;
const ty = s => d.questions.filter(q => q.type === s).length;
kids.push(P(''), table([['난이도/유형', '문항 수'], ['★☆☆ 기본', String(lv('★☆☆'))], ['★★☆ 중급', String(lv('★★☆'))], ['★★★ 심화', String(lv('★★★'))],
  ['서술형', String(ty('서술형'))], ['비교·분석형', String(ty('비교·분석형'))], ['계산·수식형', String(ty('계산·수식형'))], ['사례적용형', String(ty('사례적용형'))]], [4513, 4513]));

let md = `# 📅 ${d.dateKo} 교통기술사 일일 예상문제 50선\n생성 기준: 기출분석 + 정책동향 + 첨단모빌리티 + 미디어이슈  \n시험까지: ${d.dday}\n\n`;
for (const c of Object.keys(CATS)) {
  kids.push(new Paragraph({ children: [new PageBreak()] }),
    new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: `[카테고리 ${c}] ${CATS[c]} (5문항)`, font: F, size: 32, bold: true })] }));
  md += `## [${c}] ${CATS[c]}\n\n`;
  for (const q of d.questions.filter(x => x.cat === c)) {
    kids.push(new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 240 },
      children: [new TextRun({ text: `[${q.id}] ${q.diff} ${CATS[c]} | ${q.type}`, font: F, size: 26, bold: true, color: '1E3A5F' })] }),
      P(q.q), P(`📌 핵심 키워드: ${q.kw}`, { size: 20 }), P(`📎 참고법규: ${q.law}`, { size: 20 }),
      P(`🌐 해외기준: ${q.intl}`, { size: 20 }), P(`⭐ 출제 근거: ${q.basis}`, { size: 20 }));
    md += `### [${q.id}] ${q.diff} ${CATS[c]} | ${q.type}\n${q.q}\n\n- 📌 핵심 키워드: ${q.kw}\n- 📎 참고법규: ${q.law}\n- 🌐 해외기준: ${q.intl}\n- ⭐ 출제 근거: ${q.basis}\n\n`;
  }
}
const g = d.guide;
kids.push(new Paragraph({ children: [new PageBreak()] }),
  new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: '오늘의 학습 가이드', font: F, size: 32, bold: true })] }),
  P(`✅ 필수 정복 (★☆☆): ${g.basic} → 개념 암기 우선`), P(`⚡ 집중 학습 (★★☆): ${g.mid} → 비교표 작성 연습`),
  P(`🚀 심화 도전 (★★★): ${g.adv} → 창의적 제안 연습`), P('💡 오늘의 핵심 암기 포인트', { bold: true }),
  ...g.points.map((p, i) => P(`  ${i + 1}. ${p}`)), P(`📅 내일 예고 주제 → ${g.tomorrow}`),
  new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 360 }, children: [new TextRun({ text: '핵심 참고문헌', font: F, size: 32, bold: true })] }),
  ...d.refs.map((r, i) => P(`[${i + 1}] ${r}`, { size: 20 })));
md += `## 📊 오늘의 출제 근거 분석\n| 구분 | 반영 내용 | 문항 |\n|---|---|---|\n${d.analysis.map(r => `| ${r.join(' | ')} |`).join('\n')}\n\n`;
md += `## 📚 오늘의 학습 전략\n- ✅ 필수 정복 (★☆☆): ${g.basic}\n- ⚡ 집중 학습 (★★☆): ${g.mid}\n- 🚀 심화 도전 (★★★): ${g.adv}\n\n**💡 핵심 암기 포인트**\n${g.points.map((p, i) => `${i + 1}. ${p}`).join('\n')}\n\n**📅 내일 예고:** ${g.tomorrow}\n\n## 참고문헌\n${d.refs.map((r, i) => `${i + 1}. ${r}`).join('\n')}\n`;

const doc = new Document({
  styles: { default: { document: { run: { font: F, size: 22 } } } },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } }, children: kids }],
});
const ymd = d.date.replace(/-/g, '');
Packer.toBuffer(doc).then(b => {
  fs.writeFileSync(path.join(outDir, `교통기술사_일일예상문제_${ymd}.docx`), b);
  fs.writeFileSync(path.join(outDir, `교통기술사_일일예상문제_${ymd}.md`), md);
  console.log('ok', d.questions.length, lv('★☆☆'), lv('★★☆'), lv('★★★'), ty('서술형'), ty('비교·분석형'), ty('계산·수식형'), ty('사례적용형'));
});
