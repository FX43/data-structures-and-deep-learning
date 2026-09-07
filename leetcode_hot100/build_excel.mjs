/** Rebuild the tracking workbook from the catalog, preserving current per-problem user records. */
import fs from 'node:fs/promises';
import path from 'node:path';
import {fileURLToPath,pathToFileURL} from 'node:url';
import {createRequire} from 'node:module';
import {spawnSync} from 'node:child_process';
const root=path.dirname(fileURLToPath(import.meta.url));
const require=createRequire(import.meta.url);
const locations=process.env.CODEX_NODE_MODULES ? [process.env.CODEX_NODE_MODULES] : [root];
const {Workbook,SpreadsheetFile,FileBlob}=await import(pathToFileURL(require.resolve('@oai/artifact-tool',{paths:locations})).href);
const filename='LeetCode Hot100与高频题目刷题清单与进度追踪表.xlsx';
const catalog=JSON.parse(await fs.readFile(path.join(root,'资料','题目索引.json'),'utf8'));
const problems=catalog.problems.map(r=>({...r}));
const index=new Map(problems.map(r=>[r.id,r]));
try {
  const original=await SpreadsheetFile.importXlsx(await FileBlob.load(path.join(root,filename)));
  let sheet,isNew=true;
  try {sheet=original.worksheets.getItem('题目总表');} catch {isNew=false;sheet=original.worksheets.getItem('Hot100与高频题目');}
  if(!sheet){isNew=false;sheet=original.worksheets.getItem('Hot100与高频题目');}
  for(const row of sheet.getUsedRange().values){
    if(!Number.isInteger(row[2])) continue;
    const p=index.get(row[2]);
    if(!p) throw new Error(`现有表格中的 ${row[2]} 尚未加入题目索引，已停止重建以保护记录。`);
    const cols=isNew?[9,10,11,12,13]:[5,6,8,9,10];
    ['frequency','practice_status','rounds','notes','extra'].forEach((key,j)=>p[key]=row[cols[j]]??null);
  }
} catch(e) {
  if(e.code!=='ENOENT') throw e;
}
const w=Workbook.create();
const nav=w.worksheets.add('分类导航');
const main=w.worksheets.add('题目总表');
const notes=w.worksheets.add('使用说明');
const font={name:'Microsoft YaHei',size:11,color:'#243247'};
const blue='#234F77',pale='#EAF1F7';
function base(s,range){s.showGridLines=false;s.getRange(range).format.font=font;s.getRange(range).format.verticalAlignment='center';}
function header(s,range){s.getRange(range).format={fill:blue,font:{...font,bold:true,color:'#FFFFFF'},rowHeight:30,wrapText:true};}
function title(s,text){s.getRange('A1').values=[[text]];s.getRange('A1').format.font={...font,size:17,bold:true};s.getRange('A1').format.rowHeight=32;}
const links=[];
function link(s,cell,target,label){s.getRange(cell).values=[[label]];links.push({sheet:s===nav?'分类导航':'题目总表',cell,target,label});s.getRange(cell).format.font={...font,color:'#1765A3',underline:'single'};}
base(main,`A1:R${problems.length+4}`);
title(main,'LeetCode 题目总表');
main.getRange('A2').values=[['按官网分类排列；扩展题紧邻相关题。筛选分类或解法标签，点击题解文档后查找章节。']];
main.getRange('A3').values=[['频次、状态、刷题次数可编辑；错题补充保留原始记录。']];
const heads=['官网分类','子专题','题号','题目','题解文档','章节','笔记情况','范围','难度','频次','状态','刷题次数','解题思路与注意事项','错题补充','相关Hot100题','解法标签','官网链接','分类依据'];
main.getRange('A4:R4').values=[heads];header(main,'A4:R4');
const data=problems.map(p=>[p.category,p.subtopic,p.id,p.title,p.notebook,p.anchor,p.status,p.scope,p.difficulty,p.frequency,p.practice_status,p.rounds,p.notes,p.extra,p.related.map(String).join('、'),p.tags,p.url,p.reason]);
main.getRange(`A5:R${data.length+4}`).values=data;
main.freezePanes.freezeRows(4);main.freezePanes.freezeColumns(4);
const widths=[15,29,8,38,27,12,23,12,9,10,14,14,70,55,20,44,76,65];
widths.forEach((width,i)=>main.getRangeByIndexes(0,i,data.length+4,1).format.columnWidth=width);
main.getRange(`A5:R${data.length+4}`).format.wrapText=true;
main.getRange(`C5:C${data.length+4}`).setNumberFormat('0');
const table=main.tables.add(`A4:R${data.length+4}`,true,'ProblemIndex');
table.showFilterButton=true;
table.style='TableStyleMedium2';
const starts={};let lastCategory='';
problems.forEach((p,i)=>{
  const r=i+5;
  if(!starts[p.category])starts[p.category]=r;
  let lineCount=Math.max(...data[i].map((v,j)=>Math.ceil([...String(v??'')].reduce((a,c)=>a+(c.charCodeAt(0)>255?2:1),0)/(widths[j]-1))));
  main.getRange(`A${r}:R${r}`).format.rowHeight=Math.max(48,lineCount*17+12);
  if(p.category!==lastCategory){main.getRange(`A${r}:D${r}`).format.fill=pale;main.getRange(`A${r}:B${r}`).format.font={...font,bold:true};lastCategory=p.category;}
  main.getRange(`J${r}:L${r}`).format.fill='#FFF7DA';
  link(main,`E${r}`,p.notebook,p.notebook);
});
main.getRange(`I5:I${data.length+4}`).conditionalFormats.add('containsText',{text:'困难',format:{font:{color:'#A62838'},fill:'#FBEAEC'}});
main.getRange(`I5:I${data.length+4}`).conditionalFormats.add('containsText',{text:'中等',format:{font:{color:'#8A5916'},fill:'#FFF4D8'}});
main.getRange(`I5:I${data.length+4}`).conditionalFormats.add('containsText',{text:'简单',format:{font:{color:'#266345'},fill:'#E9F3E8'}});

base(nav,'A1:G23');title(nav,'LeetCode 分类导航');
nav.getRange('A2').values=[['按官网17类整理；题目数包含已有资料中的扩展题。点击文档打开笔记，点击查看进入总表。']];
nav.getRange('A3').values=[['分类依据：https://leetcode.cn/studyplan/top-100-liked/ （2026-09-07 核对）']];
nav.getRange('A5:G5').values=[['官网分类','本地题数','Hot100','扩展题','题解文档','总表入口','分类重点']];header(nav,'A5:G5');
const focuses={哈希:'哈希分组、集合查找',双指针:'移动零、三数之和、接雨水',滑动窗口:'异位词窗口、不定长窗口',子串:'前缀和、单调队列、覆盖窗口',普通数组:'区间、原地变换、前后缀',矩阵:'矩阵操作、搜索二维矩阵II',链表:'反转、环、归并、LRU缓存',二叉树:'遍历、路径、树形DP与一般树扩展',图论:'岛屿、拓扑排序、前缀树',回溯:'排列、子集、组合与约束',二分查找:'边界、旋转数组、中位数',栈:'辅助栈、单调栈、解码',堆:'第K大、对顶堆',贪心算法:'跳跃游戏、股票题族及状态机DP扩展',动态规划:'线性、背包、LIS、计数',多维动态规划:'回文区间、双序列、二维状态',技巧:'颜色分类、排列、判圈'};
Object.entries(catalog.categories).forEach(([cat,file],i)=>{
  const r=i+6;
  nav.getRange(`A${r}:G${r}`).values=[[cat,null,null,null,file,'查看',focuses[cat]]];
  nav.getRange(`B${r}:D${r}`).formulas=[[
    `=COUNTIF('题目总表'!$A$5:$A$${data.length+4},A${r})`,
    `=COUNTIFS('题目总表'!$A$5:$A$${data.length+4},A${r},'题目总表'!$H$5:$H$${data.length+4},"Hot100")`,
    `=COUNTIFS('题目总表'!$A$5:$A$${data.length+4},A${r},'题目总表'!$H$5:$H$${data.length+4},"扩展题")`]];
  link(nav,`E${r}`,file,file);link(nav,`F${r}`,`#'题目总表'!A${starts[cat]}`,'查看');
  nav.getRange(`A${r}:G${r}`).format.rowHeight=33;
  if(i%2===0)nav.getRange(`A${r}:G${r}`).format.fill='#F4F7FA';
});
nav.getRange('A23:D23').values=[['合计',null,null,null]];
nav.getRange('B23:D23').formulas=[['=SUM(B6:B22)','=SUM(C6:C22)','=SUM(D6:D22)']];
nav.getRange('A23:G23').format.font={...font,bold:true};nav.getRange('A23:G23').format.rowHeight=32;
[18,12,12,12,31,12,55].forEach((v,i)=>nav.getRangeByIndexes(0,i,23,1).format.columnWidth=v);
nav.getRange('A5:G23').format.wrapText=true;
const helpRows=[
 ['开始使用','先看分类导航，再从题目总表筛选官网分类、题号、范围或解法标签。'],
 ['打开题解','点击“题解文档”打开对应ipynb；“章节”中的lc-题号可在编辑器内搜索。README和笔记目录也提供章节链接。'],
 ['链接位置','保持整个leetcode文件夹的相对位置。Excel使用本地文件关联打开ipynb；编辑器不支持章节链接时，按章节标识搜索。'],
 ['维护进度','只在题目总表编辑频次、状态、刷题次数，避免分类页出现重复记录。黄色列为这些可填写项。'],
 ['分类与解法','主分类严格采用Hot100题单的分组，不等同于题目的所有算法标签。扩展题按关联题族归组。'],
 ['股票和一般树','股票状态机DP扩展随官网121股票题放在贪心算法。一般树直径扩展随543放在二叉树，子专题和解法标签明确区分。'],
 ['已有详细笔记','已找到该题的独立笔记段落。保留原代码、输出和图片，不代表原代码已通过正确性验证。'],
 ['仅有表格摘要','74、167、19未找到独立详细题解；保留表格摘要并提供待补充章节。206只有标题与摘要。'],
 ['新增索引','原表格实际101题，标题97已过期。本次补入笔记已有的235和240，最终103题，不补写其他Hot100题解。'],
 ['易混题校正','原矩阵搜索笔记适用于240，74仍保留原表的一维二分摘要；链表中误写“合并两个有序链表”的addTwoNumbers段落校正为2。'],
 ['多种解法','42的双指针、前后缀与单调栈解法集中在双指针笔记；283的两份原笔记集中在同一题下。'],
 ['原有记录','原表的摘要及无标题第11列错题记录均保留；第11列现在命名为“错题补充”。'],
 ['题解编号','原笔记局部编号保留，不等于LeetCode题号；以新增lc-题号章节及索引为准。'],
 ['官方分类来源',catalog.source],
 ['整理前Git备份','https://github.com/FX43/data-structures-and-deep-learning/tree/codex/leetcode-backup-20260907/leetcode'],
 ['备份提交','66c2daa'],
];
base(notes,`A1:B${helpRows.length+3}`);title(notes,'使用说明');
notes.getRange('A3:B3').values=[['事项','说明']];header(notes,'A3:B3');
notes.getRange(`A4:B${helpRows.length+3}`).values=helpRows;
notes.getRange(`A3:A${helpRows.length+3}`).format.columnWidth=23;
notes.getRange(`B3:B${helpRows.length+3}`).format.columnWidth=110;
notes.getRange(`A4:B${helpRows.length+3}`).format.wrapText=true;
helpRows.forEach((row,i)=>{notes.getRange(`A${i+4}:B${i+4}`).format.rowHeight=Math.max(44,Math.ceil(row[1].length/48)*19+12);});

console.log((await w.inspect({kind:'table',range:'分类导航!A5:D23',include:'values,formulas',tableMaxRows:19,tableMaxCols:4,maxChars:4000})).ndjson);
console.log((await w.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!',options:{useRegex:true,maxResults:30},summary:'final formula error scan',maxChars:2000})).ndjson);
const previewDir=process.env.CODEX_PREVIEW_DIR;
if(previewDir){
  await fs.mkdir(previewDir,{recursive:true});
  for(const [sheetName,range,key] of [['分类导航','A1:G23','navigation'],['题目总表','A4:I14','problems'],['题目总表','J4:R9','tracking'],['使用说明','A1:B19','instructions']]){
    const blob=await w.render({sheetName,range,scale:1.4});
    await fs.writeFile(path.join(previewDir,`${key}.png`),new Uint8Array(await blob.arrayBuffer()));
  }
}
// Write to a temporary sibling and replace only after export succeeds.
const temp=path.join(root,'.tracking.pending.xlsx');
await (await SpreadsheetFile.exportXlsx(w)).save(temp);
const linkfile=path.join(root,'.tracking.links.json');
await fs.writeFile(linkfile,JSON.stringify(links),'utf8');
const patch=spawnSync(process.env.CODEX_PYTHON||'python',[path.join(root,'tools','xlsx_links.py'),temp,linkfile],{stdio:'inherit'});
if(patch.status!==0)throw new Error('添加原生超链接失败；原工作表未替换。');
await fs.rename(temp,path.join(root,filename));
await fs.unlink(linkfile);
console.log(`Saved ${problems.length} problems to ${filename}`);
