const fs=require('fs'),path=require('path');
const D=require(path.resolve(process.argv[2]));const outDir=process.argv[3];
const {Document,Packer,Paragraph,TextRun,HeadingLevel,Table,TableRow,TableCell,WidthType,ShadingType,AlignmentType,PageBreak,BorderStyle}=require('docx');
const star=d=>'★'.repeat(d)+'☆'.repeat(3-d);
const all=D.categories.flatMap(c=>c.qs.map((q,i)=>({...q,id:`${c.id}-${i+1}`,cat:c.name})));
const cnt=(k)=>all.reduce((m,q)=>(m[q[k]]=(m[q[k]]||0)+1,m),{});
console.log(all.length,cnt('t'),cnt('d'));
const F="맑은 고딕";const p=(t,o={})=>new Paragraph({spacing:{line:360},...o,children:[new TextRun({text:t,font:F,size:22,...(o.run||{})})]});
const cell=(t,w,h)=>new TableCell({width:{size:w,type:WidthType.DXA},shading:h?{type:ShadingType.CLEAR,fill:"1E3A5F",color:"auto"}:undefined,children:[new Paragraph({children:[new TextRun({text:t,font:F,size:20,bold:!!h,color:h?"FFFFFF":"000000"})]})]});
const table=(rows,ws)=>new Table({width:{size:ws.reduce((a,b)=>a+b),type:WidthType.DXA},columnWidths:ws,rows:rows.map((r,i)=>new TableRow({children:r.map((t,j)=>cell(t,ws[j],i==0))}))});
const ch=[];
ch.push(new Paragraph({spacing:{before:3000},alignment:AlignmentType.CENTER,children:[new TextRun({text:"교통기술사 일일 예상문제 50선",font:F,size:48,bold:true,color:"1E3A5F"})]}));
ch.push(new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:400},children:[new TextRun({text:D.dateKo,font:F,size:28})]}));
ch.push(new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:200},children:[new TextRun({text:"기출분석 × 정책동향 × 첨단모빌리티 × 미디어이슈",font:F,size:24,italics:true})]}));
ch.push(new Paragraph({children:[new PageBreak()]}));
ch.push(new Paragraph({heading:HeadingLevel.HEADING_1,children:[new TextRun("출제 근거 분석 요약")]}));
ch.push(table([["구분","반영 내용","문항"],...D.analysis],[1800,5826,1400]));
ch.push(p(""));
ch.push(table([["유형","문항","난이도","문항"],["서술형",String(cnt('t')['서술형']),"★☆☆ 기본",String(cnt('d')[1])],["비교·분석형",String(cnt('t')['비교·분석형']),"★★☆ 중급",String(cnt('d')[2])],["계산·수식형",String(cnt('t')['계산·수식형']),"★★★ 심화",String(cnt('d')[3])],["사례적용형",String(cnt('t')['사례적용형']),"합계","50"]],[2256,2256,2256,2258]));
for(const c of D.categories){
 ch.push(new Paragraph({children:[new PageBreak()]}));
 ch.push(new Paragraph({heading:HeadingLevel.HEADING_1,children:[new TextRun(`[${c.id}] ${c.name} (5문항)`)]}));
 c.qs.forEach((q,i)=>{
  ch.push(new Paragraph({heading:HeadingLevel.HEADING_2,border:{bottom:{style:BorderStyle.SINGLE,size:6,color:"1E3A5F",space:2}},children:[new TextRun(`[${c.id}-${i+1}] ${star(q.d)} ${q.t}`)]}));
  ch.push(p(q.q));
  [["핵심 키워드",q.kw],["참고법규",q.law],["해외기준",q.intl],["출제 근거",q.basis]].forEach(([k,v])=>ch.push(new Paragraph({spacing:{line:300},indent:{left:240},children:[new TextRun({text:k+": ",font:F,size:20,bold:true,color:"1E3A5F"}),new TextRun({text:v,font:F,size:20})]})));
 });
}
ch.push(new Paragraph({children:[new PageBreak()]}));
ch.push(new Paragraph({heading:HeadingLevel.HEADING_1,children:[new TextRun("오늘의 학습 가이드")]}));
const by=d=>all.filter(q=>q.d===d).map(q=>q.id).join(', ');
ch.push(p("필수 정복(★☆☆, 개념 암기): "+by(1)));ch.push(p("집중 학습(★★☆, 비교표 작성 연습): "+by(2)));ch.push(p("심화 도전(★★★, 창의적 제안 연습): "+by(3)));
ch.push(new Paragraph({heading:HeadingLevel.HEADING_2,children:[new TextRun("핵심 암기 포인트")]}));D.memo.forEach((m,i)=>ch.push(p(`${i+1}. ${m}`)));
ch.push(p("내일 예고: "+D.tomorrow));
ch.push(new Paragraph({heading:HeadingLevel.HEADING_1,children:[new TextRun("핵심 참고문헌")]}));D.refs.forEach(r=>ch.push(p("· "+r)));
ch.push(p("※ 계산형 문항의 수치는 학습용 가상 수치입니다. 기출 출처는 한국산업인력공단 공개 자료 기준입니다.",{run:{italics:true,size:18}}));
const doc=new Document({styles:{default:{document:{run:{font:F,size:22}}},paragraphStyles:[
 {id:"Heading1",name:"Heading 1",basedOn:"Normal",next:"Normal",quickFormat:true,run:{font:F,size:32,bold:true,color:"1E3A5F"},paragraph:{spacing:{before:240,after:200},keepNext:true,outlineLevel:0}},
 {id:"Heading2",name:"Heading 2",basedOn:"Normal",next:"Normal",quickFormat:true,run:{font:F,size:26,bold:true},paragraph:{spacing:{before:300,after:120},keepNext:true,outlineLevel:1}}]},
 sections:[{properties:{page:{size:{width:11906,height:16838},margin:{top:1440,bottom:1440,left:1440,right:1440}}},children:ch}]});
const ymd=D.date.replace(/-/g,'');
Packer.toBuffer(doc).then(b=>fs.writeFileSync(path.join(outDir,`교통기술사_일일예상문제_${ymd}.docx`),b));
// markdown
let md=`# 📅 ${D.dateKo} 교통기술사 일일 예상문제 50선\n\n생성 기준: 기출분석 + 정책동향 + 첨단모빌리티 + 미디어이슈\n\n`;
for(const c of D.categories){md+=`## [${c.id}] ${c.name}\n\n`;c.qs.forEach((q,i)=>{md+=`### [${c.id}-${i+1}] ${star(q.d)} ${q.t}\n${q.q}\n\n- 📌 핵심 키워드: ${q.kw}\n- 📎 참고법규: ${q.law}\n- 🌐 해외기준: ${q.intl}\n- ⭐ 출제 근거: ${q.basis}\n\n`;});}
md+=`## 📊 출제 근거 분석\n\n| 구분 | 반영 내용 | 문항 |\n|---|---|---|\n`+D.analysis.map(r=>`| ${r.join(' | ')} |`).join('\n')+`\n\n## 📚 학습 전략\n- ✅ 필수 정복(★☆☆): ${by(1)}\n- ⚡ 집중 학습(★★☆): ${by(2)}\n- 🚀 심화 도전(★★★): ${by(3)}\n\n### 💡 핵심 암기 포인트\n`+D.memo.map((m,i)=>`${i+1}. ${m}`).join('\n')+`\n\n### 📅 내일 예고\n→ ${D.tomorrow}\n\n### 참고문헌\n`+D.refs.map(r=>`- ${r}`).join('\n')+`\n\n> 계산형 수치는 학습용 가상 수치입니다.\n`;
fs.writeFileSync(path.join(outDir,`교통기술사_일일예상문제_${ymd}.md`),md);
