/** Original-format tracker. Preserve per-problem notes and count appended practice marks. */
import fs from 'node:fs/promises';
import path from 'node:path';
import {fileURLToPath,pathToFileURL} from 'node:url';
import {createRequire} from 'node:module';
import {spawnSync} from 'node:child_process';
const root=path.dirname(fileURLToPath(import.meta.url));
const require=createRequire(import.meta.url);
const {Workbook,SpreadsheetFile,FileBlob}=await import(pathToFileURL(require.resolve('@oai/artifact-tool',{paths:process.env.CODEX_NODE_MODULES?[process.env.CODEX_NODE_MODULES]:[root]})).href);
const filename='LeetCode Hot100与高频题目刷题清单与进度追踪表.xlsx';
const input=JSON.parse(await fs.readFile(path.join(root,'资料/原表样式与记录.json'),'utf8'));
const records=structuredClone(input.records);
// Read the latest saved user records so rebuilding never resets practice or reminders.
const current=await SpreadsheetFile.importXlsx(await FileBlob.load(path.join(root,filename)));
let currentSheet;
try{currentSheet=current.worksheets.getItem(input.sheet);}catch{}
if(currentSheet){
  const byid=new Map(records.map(r=>[r.id,r]));
  const twoTags=currentSheet.getRange('B3').values[0][0]==='小标签';
  for(const raw of currentSheet.getUsedRange().values){
    const row=twoTags?[raw[0],...raw.slice(2)]:raw;
    if(!Number.isInteger(row[2]))continue;
    const record=byid.get(row[2]);
    if(!record)throw new Error(`题号${row[2]}尚未登记，停止重建以保留新增记录。`);
    for(const col of [3,4,5,6,7,8,10,11])record.values[col]=row[col]??null;
  }
}
const w=Workbook.create();const s=w.worksheets.add(input.sheet);
const last=records.length+3;
const widths=[17,28,...input.widths.slice(1)];
widths.forEach((width,col)=>s.getRangeByIndexes(0,col,last,1).format.columnWidth=width);
s.getRange('A1:M1').merge();s.getRange('A1:M1').format=input.title_style;
s.getRange('A1').values=[[`LeetCode Hot 100 与高频题目总数：${records.length}`]];
s.getRange('A1:M1').format.rowHeight=34;
s.getRange('A2:M2').merge();s.getRange('A2:M2').format=input.intro_style;
s.getRange('A2').values=[['A列大标签采用Hot100官网分类，B列小标签说明具体题型。递归等解法按题型放入小标签。J列每做一次就在末尾追加√或×，K列自动统计次数（兼容✓）；双击或F2编辑以保留旧记录。原笔记、错题补充全部保留。点击题目打开ipynb，按lc-题号搜索。']];
s.getRange('A2:M2').format.rowHeight=46;
s.getRange('A3:M3').values=[['大标签','小标签',...input.headers.slice(1)]];s.getRange('A3:M3').format=input.header_style;
s.getRange('A3:M3').format.rowHeight=48;
s.freezePanes.freezeRows(3);s.freezePanes.freezeColumns(2);
const links=[];const groupRanges=[];let groupStart=4,previousGroup='';
const majorRanges=[];let majorStart=4,previousMajor='';
function textUnits(value){return String(value??'').split('\n').map(line=>[...line].reduce((n,c)=>n+(c.charCodeAt(0)>255?2:1),0));}
function countFormula(row){return `=LEN(J${row})-LEN(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(J${row},"√",""),"×",""),"✓",""))`;}
for(const [i,record] of records.entries()){
  const row=i+4;const group=record.group.join('\n');
  if(i&&group!==previousGroup){groupRanges.push([groupStart,row-1,previousGroup]);groupStart=row;}
  previousGroup=group;
  if(i&&record.group[0]!==previousMajor){majorRanges.push([majorStart,row-1,previousMajor]);majorStart=row;}
  previousMajor=record.group[0];
  const values=[record.group[0],record.group[1],...record.values.slice(1)];
  const styles=[record.styles[0],record.styles[0],...record.styles.slice(1)];
  s.getRange(`A${row}:M${row}`).values=[values];
  styles.forEach((format,j)=>s.getCell(row-1,j).format=format);
  s.getRange(`E${row}`).format.wrapText=true;
  s.getRange(`J${row}`).setNumberFormat('@');
  s.getRange(`K${row}`).formulas=[[countFormula(row)]];s.getRange(`K${row}`).setNumberFormat('0');
  s.getRange(`K${row}`).format.font={name:'微软雅黑',size:11,bold:true,color:'#1F4E79'};
  // Match the existing notes style; increase only the row height needed to show complete text.
  let height=record.height;
  for(const col of [0,1,4,8,9,11,12]){
    const lines=textUnits(values[col]).reduce((n,len)=>n+Math.max(1,Math.ceil(len/(widths[col]-2))),0);
    const fontSize=styles[col].font.size;
    height=Math.max(height,lines*(fontSize*1.4)+8);
  }
  s.getRange(`A${row}:M${row}`).format.rowHeight=Math.min(409,Math.max(36,height));
  links.push({sheet:input.sheet,cell:`E${row}`,target:record.notebook,label:record.values[3],tooltip:`${record.notebook} / ${record.anchor}`});
  links.push({sheet:input.sheet,cell:`I${row}`,target:record.url,label:record.values[7]});
}
groupRanges.push([groupStart,last,previousGroup]);
majorRanges.push([majorStart,last,previousMajor]);
for(const [col,ranges] of [['A',majorRanges],['B',groupRanges]])for(const [start,end,label] of ranges){
  s.getRange(`${col}${start}:${col}${end}`).clear({applyTo:'contents'});
  if(end>start)s.getRange(`${col}${start}:${col}${end}`).merge();
  s.getRange(`${col}${start}`).values=[[col==='A'?label:label.split('\n').slice(1).join('\n')]];
  s.getRange(`${col}${start}:${col}${end}`).format=input.records[0].styles[0];
  if(col==='B')s.getRange(`${col}${start}:${col}${end}`).format.fill='#F2F6FA';
}
s.getRange(`J4:J${last}`).dataValidation={
  allowBlank:true,
  rule:{type:'custom',formula1:'=LEN(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(J4,"√",""),"×",""),"✓",""))=0'},
  prompt:{title:'追加本次结果',message:'双击或F2编辑，在末尾追加√或×。例如√×√表示3次；不要覆盖已有记录。'},
  errorAlert:{style:'stop',title:'只填写做题结果',message:'请只填写√或×（兼容✓）。每个符号代表做了一次，例如√×√。'}
};
w.recalculate();
// Meaningful recalculation checks; restore every temporary value before saving.
const tests=[];
for(const row of [4,Math.floor((4+last)/2),last]){
  const saved=s.getRange(`J${row}`).values[0][0];
  for(const [marks,expected] of [[null,0],['√',1],['×',1],['√×√',3],['✓×',2],['√×✓√×',5]]){
    s.getRange(`J${row}`).values=[[marks]];w.recalculate();
    const actual=s.getRange(`K${row}`).values[0][0];
    if(actual!==expected)throw new Error(`Counter test failed at ${row}: ${marks}: ${actual}`);
    tests.push({row,marks,expected,actual});
  }
  s.getRange(`J${row}`).values=[[saved]];
}
w.recalculate();
console.log(JSON.stringify({problems:records.length,majorGroups:majorRanges.length,minorGroups:groupRanges.length,counterChecks:tests.length}));
console.log((await w.inspect({kind:'table',range:`${input.sheet}!J3:M7`,include:'values,formulas',tableMaxRows:5,tableMaxCols:4,maxChars:2000})).ndjson);
console.log((await w.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!',options:{useRegex:true,maxResults:20},maxChars:1000})).ndjson);
if(process.env.CODEX_PREVIEW_DIR){
  const out=process.env.CODEX_PREVIEW_DIR;
  for(const [range,name] of [['A3:F16','two-tags'],['I3:M9','two-tags-notes']]){
    const blob=await w.render({sheetName:input.sheet,range,scale:1.5});
    await fs.writeFile(path.join(out,`${name}.png`),new Uint8Array(await blob.arrayBuffer()));
  }
  const saved=s.getRange('J4').values[0][0];s.getRange('J4').values=[['√×√']];w.recalculate();
  const example=await w.render({sheetName:input.sheet,range:'J3:K5',scale:2});
  await fs.writeFile(path.join(out,'practice-example.png'),new Uint8Array(await example.arrayBuffer()));
  s.getRange('J4').values=[[saved]];w.recalculate();
  await fs.writeFile(path.join(out,'counter-verification.json'),JSON.stringify(tests,null,2));
}
const temp=path.join(root,'.tracking.pending.xlsx');
await (await SpreadsheetFile.exportXlsx(w)).save(temp);
const linkfile=path.join(root,'.tracking.links.json');await fs.writeFile(linkfile,JSON.stringify(links));
const patch=spawnSync(process.env.CODEX_PYTHON||'python',[path.join(root,'tools/xlsx_links.py'),temp,linkfile,'two-tags'],{stdio:'inherit'});
if(patch.status!==0)throw new Error('Excel链接处理失败，原文件未替换。');
await fs.rename(temp,path.join(root,filename));await fs.unlink(linkfile);
console.log(`Saved original-format tracker with ${records.length} problems.`);
