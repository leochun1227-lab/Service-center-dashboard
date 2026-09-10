import fs from 'node:fs/promises';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {Workbook,SpreadsheetFile} from '@oai/artifact-tool';
const dir=path.dirname(fileURLToPath(import.meta.url));
const data=JSON.parse(await fs.readFile(path.join(dir,'comparison_data.json'),'utf8'));
const months=new Set(data.months);
const records=data.dashboard.filter(r=>months.has(r[5])||months.has(r[8])).sort((a,b)=>a[1].localeCompare(b[1])||a[5].localeCompare(b[5])||a[0].localeCompare(b[0]));
const dt=x=>typeof x==='string'&&/^\d{4}-\d{2}-\d{2}$/.test(x)?new Date(x+'T00:00:00Z'):x;
const wb=Workbook.create();
const s=wb.worksheets.add('Summary'),t=wb.worksheets.add('Ticket Details');
const num='#,##0.00;(#,##0.00);0.00';
const end=records.length+1;
for(const [sh,range] of [[s,'A1:D93'],[t,`A1:L${end}`]]){sh.showGridLines=false;sh.getRange(range).format.font={name:'Arial',size:10,color:'#000000'};sh.getRange(range).format.rowHeight=20;sh.getRange(range).format.verticalAlignment='center';}
t.getRange('A1:L1').values=[['TicketID','Dealer','Ticket Type','Status','CreatedOn','Created Month','TotalLabourHours','Invoice Number','Billing Date','Invoice Month','Invoice Amount','Invoiced']];
t.getRange(`A2:L${end}`).values=records.map(r=>[r[0],r[1],r[2],r[3],dt(r[4]),dt(r[5]),r[9],r[6],dt(r[7]),dt(r[8]),r[10],r[11]]);
for(const c of ['E','I']){t.getRange(`${c}2:${c}${end}`).setNumberFormat('dd-mmm-yyyy');t.getRange(`${c}2:${c}${end}`).format.horizontalAlignment='center';}
for(const c of ['F','J']){t.getRange(`${c}2:${c}${end}`).setNumberFormat('mmm yyyy');t.getRange(`${c}2:${c}${end}`).format.horizontalAlignment='center';}
for(const c of ['G','K'])t.getRange(`${c}2:${c}${end}`).setNumberFormat(num);
for(const c of ['A','H'])t.getRange(`${c}2:${c}${end}`).setNumberFormat('@');
const widths=[13,18,19,24,16,17,23,23,16,17,21,13];widths.forEach((w,i)=>t.getRange(`${String.fromCharCode(65+i)}1:${String.fromCharCode(65+i)}${end}`).format.columnWidth=w);
const table=t.tables.add(`A1:L${end}`,true,'Tickets');table.style='TableStyleLight1';t.getRange(`A1:L${end}`).format.fill='#FFFFFF';t.getRange('A1:L1').format.font.bold=true;t.getRange('A1:L1').format.wrapText=true;t.getRange('A1:L1').format.rowHeight=30;t.getRange('A1:L1').format.borders={bottom:{style:'thin',color:'#000000'}};t.freezePanes.freezeRows(1);t.freezePanes.freezeColumns(2);
s.getRange('A1:A93').format.columnWidth=24;s.getRange('B1:B93').format.columnWidth=23;s.getRange('C1:C93').format.columnWidth=22;s.getRange('D1:D93').format.columnWidth=23;
const ref=c=>`'Ticket Details'!$${c}$2:$${c}$${end}`;
const locations=['Perth','Traralgon','Launceston','Geelong','Frankston','Other'];
const groups=[{offset:0,title:'All Service',type:null},{offset:32,title:'Repair Only',type:'Repair ticket'},{offset:64,title:'PDI Only',type:'PDI'}];
if(records.some(r=>!['Repair ticket','PDI'].includes(r[2])))throw Error('Unexpected ticket type');
for(const group of groups){
const o=group.offset,criteria=group.type?`,${ref('C')},"${group.type}"`:'';
s.getRange(`A${o+1}`).values=[[group.title]];s.getRange(`A${o+1}`).format.font.bold=true;s.getRange(`A${o+2}:B${o+2}`).values=[['Year',2026]];
s.getRange(`A${o+4}:D${o+4}`).values=[['Row Labels','Total Labour Hours','Invoiced Tickets','Invoice Amount']];s.getRange(`A${o+4}:D${o+4}`).format.font.bold=true;s.getRange(`A${o+4}:D${o+4}`).format.rowHeight=24;s.getRange(`A${o+4}:D${o+4}`).format.borders={bottom:{style:'thin',color:'#000000'}};
for(let k=0;k<locations.length;k++){
 const h=o+5+k*4;s.getRange(`A${h}`).values=[[locations[k]]];s.getRange(`A${h}`).format.font.bold=true;
 for(let i=0;i<3;i++){
  const r=h+i+1;s.getRange(`A${r}`).values=[[dt(data.months[i])]];s.getRange(`A${r}`).setNumberFormat('    mmm');
  s.getRange(`B${r}:D${r}`).formulas=[[
   `=SUMIFS(${ref('G')},${ref('B')},$A$${h},${ref('F')},$A${r}${criteria})`,
   `=COUNTIFS(${ref('B')},$A$${h},${ref('J')},$A${r},${ref('L')},"Yes"${criteria})`,
   `=SUMIFS(${ref('K')},${ref('B')},$A$${h},${ref('J')},$A${r},${ref('L')},"Yes",${ref('H')},"<>"${criteria})`]];
 }
}
s.getRange(`A${o+5}:A${o+28}`).format.horizontalAlignment='left';
s.getRange(`A${o+29}`).values=[['Grand Total']];for(const c of ['B','C','D'])s.getRange(`${c}${o+29}`).formulas=[[`=SUM(${c}${o+5}:${c}${o+28})`]];
s.getRange(`A${o+29}:D${o+29}`).format.font.bold=true;s.getRange(`A${o+29}:D${o+29}`).format.borders={top:{style:'thin',color:'#000000'}};
s.getRange(`B${o+5}:B${o+29}`).setNumberFormat(num);s.getRange(`C${o+5}:C${o+29}`).setNumberFormat('#,##0');s.getRange(`D${o+5}:D${o+29}`).setNumberFormat(num);
}
wb.recalculate();
for(const group of groups){
const scoped=records.filter(x=>!group.type||x[2]===group.type);
for(let k=0;k<locations.length;k++)for(let i=0;i<3;i++){
 const m=data.months[i],location=locations[k],r=group.offset+6+k*4+i;
 const dh=scoped.filter(x=>x[1]===location&&x[5]===m).reduce((a,x)=>a+x[9],0);
 const di=scoped.filter(x=>x[1]===location&&x[8]===m&&x[11]==='Yes');
 const expected=[dh,di.length,di.reduce((a,x)=>a+x[12],0)];const values=s.getRange(`B${r}:D${r}`).values[0];
 expected.forEach((v,j)=>{if(Math.abs(v-values[j])>1e-7)throw Error(`Mismatch ${location} ${m} ${j}`)});
}
console.log(group.title,s.getRange(`B${group.offset+29}:D${group.offset+29}`).values);
}
for(let row=5;row<=29;row++)for(const c of ['B','C','D']){
 const value=r=>s.getRange(`${c}${r}`).values[0][0]||0;
 if(Math.abs(value(row)-value(row+32)-value(row+64))>1e-7)throw Error(`Service split mismatch ${c}${row}`);
}
console.log((await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!',options:{useRegex:true,maxResults:10},maxChars:1000})).ndjson);
console.log('DETAIL ROWS',records.length,'TOTAL',s.getRange('A29:D29').values);
const outfile=path.join(dir,'Service_Monthly_Totals_By_Service.xlsx');await(await SpreadsheetFile.exportXlsx(wb)).save(outfile);
for(const [sh,range,n] of [[s,'A1:D29','service_all'],[s,'A33:D61','service_repair'],[s,'A65:D93','service_pdi'],[t,'A1:L10','service_details']]){const img=await wb.render({sheetName:sh.name,range,scale:1.5,format:'png'});await fs.writeFile(path.join(dir,n+'.png'),new Uint8Array(await img.arrayBuffer()));}
console.log(outfile);
