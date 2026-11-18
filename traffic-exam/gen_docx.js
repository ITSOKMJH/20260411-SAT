// Usage: node gen_docx.js questions_YYYYMMDD.json
const fs=require('fs');const path=require('path');
const D=require('docx');const {Document,Packer,Paragraph,TextRun,Table,TableRow,TableCell,HeadingLevel,AlignmentType,WidthType,ShadingType,PageBreak,BorderStyle}=D;
const data=JSON.parse(fs.readFileSync(process.argv[2],'utf8'));
const F='맑은 고딕';const ymd=data.date.replace(/-/g,'');const [y,m,d]=data.date.split('-');
const all=data.cats.flatMap(c=>c.q.map((q,i)=>({id:`${c.id}-${i+1}`,cat:c.name,diff:q[0],type:q[1],text:q[2],kw:q[3],law:q[4],intl:q[5],basis:q[6]})));
const cnt=(k)=>all.reduce((a,x)=>(a[x[k]]=(a[x[k]]||0)+1,a),{});
const tc=cnt('type'),dc=cnt('diff');
if(all.length!==50)throw new Error('need 50');
console.log(tc,dc);
const P=(t,o={})=>new Paragraph({spacing:{line:360,after:80},alignment:o.align,children:[new TextRun({text:t,font:F,size:o.size||22,bold:o.bold,color:o.color})]});
const H1=t=>new Paragraph({heading:HeadingLevel.HEADING_1,spacing:{before:240,after:120},children:[new TextRun({text:t,font:F,size:32,bold:true,color:'1E3A5F'})]});
const H2=t=>new Paragraph({heading:HeadingLevel.HEADING_2,spacing:{before:200,after:80},children:[new TextRun({text:t,font:F,size:26,bold:true})]});
const cell=(t,h,w)=>new TableCell({width:{size:w,type:WidthType.DXA},shading:h?{type:ShadingType.CLEAR,fill:'1E3A5F',color:'auto'}:undefined,margins:{top:60,bottom:60,left:100,right:100},children:[new Paragraph({children:[new TextRun({text:t,font:F,size:20,bold:h,color:h?'FFFFFF':undefined})]})]});
const table=(head,rows,ws)=>new Table({width:{size:9026,type:WidthType.DXA},columnWidths:ws,rows:[new TableRow({tableHeader:true,children:head.map((h,i)=>cell(h,true,ws[i]))}),...rows.map(r=>new TableRow({children:r.map((t,i)=>cell(String(t),false,ws[i]))}))]});
const kids=[];
kids.push(new Paragraph({spacing:{before:3000}}),P('교통기술사 일일 예상문제 50선',{size:48,bold:true,align:AlignmentType.CENTER,color:'1E3A5F'}),P(`${y}년 ${+m}월 ${+d}일`,{size:32,align:AlignmentType.CENTER}),P('“기출분석 × 정책동향 × 첨단모빌리티”',{size:24,align:AlignmentType.CENTER}),P(`문항 구성: 서술형 ${tc['서술형']} · 비교·분석형 ${tc['비교·분석형']} · 계산·수식형 ${tc['계산·수식형']} · 사례적용형 ${tc['사례적용형']}`,{align:AlignmentType.CENTER}),P(`난이도: ★☆☆ ${dc['★☆☆']} · ★★☆ ${dc['★★☆']} · ★★★ ${dc['★★★']}`,{align:AlignmentType.CENTER}),new Paragraph({children:[new PageBreak()]}));
kids.push(H1('출제 근거 분석 요약'),table(['구분','반영 내용','문항'],data.analysis.map(a=>[a[0],a[1],a[2]]),[1800,5926,1300]),P(''),table(['카테고리','문항수','세부 주제'],data.cats.map(c=>[`[${c.id}] ${c.name}`,'5',c.q.map(q=>q[3].split(',')[0]).join(' / ')]),[2400,900,5726]),new Paragraph({children:[new PageBreak()]}));
for(const c of data.cats){kids.push(H1(`[카테고리 ${c.id}] ${c.name} (5문항)`));
 c.q.forEach((q,i)=>{kids.push(H2(`[${c.id}-${i+1}] ${q[0]} ${c.name} | ${q[1]}`),P(q[2]),P(`📌 핵심 키워드: ${q[3]}`,{size:20}),P(`📎 참고법규: ${q[4]}`,{size:20}),P(`🌐 해외기준: ${q[5]}`,{size:20}),P(`⭐ 출제 근거: ${q[6]}`,{size:20,color:'8B0000'}))});}
kids.push(new Paragraph({children:[new PageBreak()]}),H1('오늘의 학습 가이드'),...data.guide.map(g=>P(g)),H1('핵심 참고문헌'),...data.refs.map(r=>P('• '+r,{size:20})));
const doc=new Document({styles:{default:{document:{run:{font:F,size:22}}}},sections:[{properties:{page:{size:{width:11906,height:16838},margin:{top:1440,bottom:1440,left:1440,right:1440}}},children:kids}]});
const out=path.join(path.dirname(process.argv[2]),`교통기술사_일일예상문제_${ymd}.docx`);
Packer.toBuffer(doc).then(b=>{fs.writeFileSync(out,b);console.log('wrote',out)});
// markdown
let md=`# 📅 ${y}년 ${+m}월 ${+d}일 교통기술사 일일 예상문제 50선\n\n생성 기준: 기출분석 + 정책동향 + 첨단모빌리티 + 미디어이슈\n\n`;
for(const c of data.cats){md+=`## [${c.id}] ${c.name}\n\n`;c.q.forEach((q,i)=>{md+=`### [${c.id}-${i+1}] ${q[0]} | ${q[1]}\n${q[2]}\n\n- 📌 키워드: ${q[3]}\n- 📎 참고법규: ${q[4]}\n- 🌐 해외기준: ${q[5]}\n- ⭐ 출제 근거: ${q[6]}\n\n`})}
md+=`## 📊 출제 근거 분석\n\n| 구분 | 내용 | 문항 |\n|---|---|---|\n`+data.analysis.map(a=>`| ${a.join(' | ')} |`).join('\n')+`\n\n## 📚 학습 가이드\n\n`+data.guide.map(g=>`- ${g}`).join('\n')+`\n\n## 참고문헌\n\n`+data.refs.map(r=>`- ${r}`).join('\n')+'\n';
fs.writeFileSync(path.join(path.dirname(process.argv[2]),`교통기술사_일일예상문제_${ymd}.md`),md);
