// 사용법: node traffic-exam/scripts/generate.js traffic-exam/data/YYYY-MM-DD.js
// 산출물: traffic-exam/output/교통기술사_일일예상문제_YYYYMMDD.docx 및 .md
const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ShadingType,
  AlignmentType, HeadingLevel, PageBreak, BorderStyle,
} = require('docx');

const data = require(path.resolve(process.argv[2]));
const outDir = path.join(__dirname, '..', 'output');
fs.mkdirSync(outDir, { recursive: true });
const ymd = data.date.replace(/-/g, '');
const stars = (n) => '★'.repeat(n) + '☆'.repeat(3 - n);
const dday = Math.round((new Date(data.examDate) - new Date(data.date)) / 86400000);

// ── 검증: 50문항, 카테고리별 5, 유형·난이도 배분
const qs = data.questions;
const count = (f) => qs.filter(f).length;
const typeCnt = {}; const lvCnt = {};
qs.forEach((q) => { typeCnt[q.type] = (typeCnt[q.type] || 0) + 1; lvCnt[q.lv] = (lvCnt[q.lv] || 0) + 1; });
if (qs.length !== 50) throw new Error('문항 수 ' + qs.length);
Object.keys(data.categories).forEach((c) => { if (count((q) => q.id[0] === c) !== 5) throw new Error('카테고리 ' + c); });
console.log('유형', typeCnt, '난이도', lvCnt);

const FONT = '맑은 고딕';
const W = 9026; // A4 - 여백
const run = (t, o = {}) => new TextRun({ text: t, font: FONT, size: 22, ...o });
const p = (t, o = {}, po = {}) => new Paragraph({ spacing: { line: 360 }, ...po, children: [run(t, o)] });
const cell = (t, w, head = false) => new TableCell({
  width: { size: w, type: WidthType.DXA },
  shading: head ? { type: ShadingType.CLEAR, fill: '1E3A5F', color: 'auto' } : undefined,
  margins: { top: 60, bottom: 60, left: 100, right: 100 },
  children: [new Paragraph({ children: [run(t, head ? { bold: true, color: 'FFFFFF', size: 20 } : { size: 20 })] })],
});
const table = (widths, rows) => new Table({
  width: { size: W, type: WidthType.DXA }, columnWidths: widths,
  rows: rows.map((r, i) => new TableRow({ tableHeader: i === 0, children: r.map((t, j) => cell(String(t), widths[j], i === 0)) })),
});
const h1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 240, after: 120 }, children: [run(t, { bold: true, size: 32, color: '1E3A5F' })] });
const h2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 200, after: 80 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '1E3A5F', space: 2 } }, children: [run(t, { bold: true, size: 26 })] });
const brk = () => new Paragraph({ children: [new PageBreak()] });

const body = [];
// 표지
body.push(new Paragraph({ spacing: { before: 3000 } }));
body.push(p('교통기술사 일일 예상문제 50선', { bold: true, size: 48, color: '1E3A5F' }, { alignment: AlignmentType.CENTER }));
body.push(p(data.dateKo, { size: 32 }, { alignment: AlignmentType.CENTER, spacing: { before: 400, line: 360 } }));
body.push(p('"기출분석 × 정책동향 × 첨단모빌리티 × 미디어이슈"', { italics: true, size: 26 }, { alignment: AlignmentType.CENTER, spacing: { before: 300, line: 360 } }));
body.push(p(`목표 시험: ${data.examTarget}  |  D-${dday}`, { size: 22 }, { alignment: AlignmentType.CENTER, spacing: { before: 600, line: 360 } }));
body.push(brk());

// 출제 근거
body.push(h1('출제 근거 분석 요약'));
body.push(table([1800, 5926, 1300], [['구분', '반영 내용', '문항 수'], ...data.analysis.map((a) => [a[0], a[1], a[2] + '문항'])]));
body.push(p(''));
body.push(table([3000, 3013, 3013], [['문제 유형', '문항 수', '난이도 분포'],
  ['서술형', (typeCnt['서술형'] || 0) + '문항', `★☆☆ 기본 ${lvCnt[1]}문항`],
  ['비교·분석형', (typeCnt['비교·분석형'] || 0) + '문항', `★★☆ 중급 ${lvCnt[2]}문항`],
  ['계산·수식형', (typeCnt['계산·수식형'] || 0) + '문항', `★★★ 심화 ${lvCnt[3]}문항`],
  ['사례적용형', (typeCnt['사례적용형'] || 0) + '문항', '합계 50문항']]));
body.push(p('※ 기출 원문 폴더(로컬 PC 경로)는 클라우드 실행환경에서 접근할 수 없어, 스킬 내장 기출 패턴(110~138회 분석)과 웹 최신 동향을 기준으로 작성했습니다. 계산문제의 수치는 가상 수치입니다.', { size: 18, color: '666666' }));

// 문항
Object.entries(data.categories).forEach(([c, name]) => {
  body.push(brk());
  body.push(h1(`[카테고리 ${c}] ${name} (5문항)`));
  qs.filter((q) => q.id[0] === c).forEach((q) => {
    body.push(h2(`[${q.id}] ${stars(q.lv)}  ${q.type}`));
    body.push(p(q.q));
    [['📌 핵심 키워드', q.kw], ['📎 참고법규', q.law], ['🌐 해외기준', q.intl], ['⭐ 출제 근거', q.basis]].forEach(([k, v]) => {
      body.push(new Paragraph({ spacing: { line: 300 }, indent: { left: 200 }, children: [run(k + ': ', { bold: true, size: 20 }), run(v, { size: 20 })] }));
    });
  });
});

// 학습 가이드
const ids = (lv) => qs.filter((q) => q.lv === lv).map((q) => q.id).join(', ');
body.push(brk());
body.push(h1('오늘의 학습 가이드'));
body.push(p('✅ 필수 정복 (★☆☆): ' + ids(1) + ' → 개념 암기 우선'));
body.push(p('⚡ 집중 학습 (★★☆): ' + ids(2) + ' → 비교표 작성 연습'));
body.push(p('🚀 심화 도전 (★★★): ' + ids(3) + ' → 창의적 제안 연습'));
body.push(p('💡 오늘의 핵심 암기 포인트', { bold: true }));
data.memo.forEach((m, i) => body.push(p(`  ${i + 1}. ${m}`)));
body.push(p('📅 내일 예고: ' + data.tomorrow, { bold: true }));
body.push(h1('핵심 참고문헌 및 출처'));
data.refs.forEach((r) => body.push(p('• ' + r, { size: 20 })));
data.sources.forEach((s) => body.push(p(`• ${s[0]}: ${s[1]}`, { size: 18, color: '1E3A5F' })));

const doc = new Document({
  styles: { default: { document: { run: { font: FONT, size: 22 } } } },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } }, children: body }],
});
const docxPath = path.join(outDir, `교통기술사_일일예상문제_${ymd}.docx`);
Packer.toBuffer(doc).then((b) => { fs.writeFileSync(docxPath, b); console.log('DOCX:', docxPath); });

// Markdown
const md = [];
md.push(`# 📅 ${data.dateKo} 교통기술사 일일 예상문제 50선`, '', `생성 기준: 기출분석 + 정책동향 + 첨단모빌리티 + 미디어이슈  `, `목표 시험: ${data.examTarget} — **D-${dday}**`, '');
Object.entries(data.categories).forEach(([c, name]) => {
  md.push(`## [${c}] ${name}`, '');
  qs.filter((q) => q.id[0] === c).forEach((q) => {
    md.push(`### [${q.id}] ${stars(q.lv)} ${q.type}`, q.q, '', `- 📌 **키워드**: ${q.kw}`, `- 📎 **참고법규**: ${q.law}`, `- 🌐 **해외기준**: ${q.intl}`, `- ⭐ **출제 근거**: ${q.basis}`, '');
  });
});
md.push('## 📊 출제 근거 분석', '', '| 구분 | 반영 내용 | 문항 |', '|---|---|---|', ...data.analysis.map((a) => `| ${a[0]} | ${a[1]} | ${a[2]} |`), '');
md.push('## 📚 학습 가이드', `- ✅ 필수(★☆☆): ${ids(1)}`, `- ⚡ 집중(★★☆): ${ids(2)}`, `- 🚀 심화(★★★): ${ids(3)}`, '', '**💡 핵심 암기**', ...data.memo.map((m, i) => `${i + 1}. ${m}`), '', `**📅 내일 예고**: ${data.tomorrow}`, '');
md.push('## 출처', ...data.sources.map((s) => `- [${s[0]}](${s[1]})`), '');
fs.writeFileSync(path.join(outDir, `교통기술사_일일예상문제_${ymd}.md`), md.join('\n'));
