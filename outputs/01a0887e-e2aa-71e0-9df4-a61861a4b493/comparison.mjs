import fs from 'node:fs/promises';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {Workbook,SpreadsheetFile} from '@oai/artifact-tool';
const dir=path.dirname(fileURLToPath(import.meta.url));
const d=JSON.parse(await fs.readFile(path.join(dir,'comparison_data.json'),'utf8'));
const wb=Workbook.create();
const sum=wb.worksheets.add('Summary'),match=wb.worksheets.add('Matched Tickets'),dash=wb.worksheets.add('Dashboard Tickets'),report=wb.worksheets.add('Report Details'),inv=wb.worksheets.add('Invoice Lines');
const col=n=>{let s='';for(n++;n;n=Math.floor((n-1)/26))s=String.fromCharCode(65+(n-1)%26)+s;return s};
const date=x=>typeof x==='string'&&/^\d{4}-\d{2}-\d{2}$/.test(x)?new Date(x+'T00:00:00Z'):x;
const num='#,##0.00;(#,##0.00);0.00';
const pct='0.0%;(0.0%);0.0%';
function style(s,n,c){s.showGridLines=false;const r=s.getRange(`A1:${c}${n}`);r.format.font={name:'Arial',size:10,color:'#000000'};r.format.rowHeight=20;r.format.verticalAlignment='center';}
function heading(s,r,end){const a=s.getRange(`A${r}:${end}${r}`);a.format.font.bold=true;a.format.wrapText=true;a.format.rowHeight=32;a.format.horizontalAlignment='center';a.format.borders={bottom:{style:'thin',color:'#000000'}};}
function total(s,r,end){s.getRange(`A${r}:${end}${r}`).format.font.bold=true;s.getRange(`A${r}:${end}${r}`).format.borders={top:{style:'thin',color:'#000000'}};}
function table(s,rows,headers,widths,name,notes){const n=rows.length+4,c=col(headers.length-1);style(s,n,c);s.getRange('A1').values=[[s.name]];s.getRange('A1').format.font.bold=true;s.getRange('A2').values=[[notes]];s.getRange(`A4:${c}4`).values=[headers];s.getRange(`A5:${c}${n}`).values=rows.map(r=>r.map(date));widths.forEach((w,i)=>s.getRange(`${col(i)}1:${col(i)}${n}`).format.columnWidth=w);const t=s.tables.add(`A4:${c}${n}`,true,name);t.style='TableStyleLight1';s.getRange(`A4:${c}${n}`).format.fill='#FFFFFF';heading(s,4,c);s.freezePanes.freezeRows(4);s.freezePanes.freezeColumns(2);return n;}
function format(s,n,cols,f){for(const c of cols){s.getRange(`${c}5:${c}${n}`).setNumberFormat(f);if(f.includes('mmm'))s.getRange(`${c}5:${c}${n}`).format.horizontalAlignment='center';}}
const dn=table(dash,d.dashboard,['TicketID','Location','Ticket Type','Status','CreatedOn','Created Month','Invoice Number','Billing Date','Invoice Month','TotalLabourHours','Invoice ex GST','In Invoice Total','Included Invoice','Common Location','Matched Order','TECO Month','Source Sheet','Source Row'],[13,17,19,23,15,16,21,15,16,22,20,19,20,20,19,17,17,13],'DashboardTickets','Source: current dashboard-data.js and ticket workbook. All Ticket Types, All Invoices.');
format(dash,dn,['E','H'],'dd-mmm-yyyy');format(dash,dn,['F','I','P'],'mmm yyyy');format(dash,dn,['J','K','M'],num);format(dash,dn,['A','G','O'],'@');
const rn=table(report,d.report,['Order','Confirmation','Work Center','Location','AR SO','TECO Date','TECO Month','Posting Date','Labour Hours','Invoice Amount','All Service','Retail Service','Matched TicketID','Match','Source Row'],[15,17,17,18,17,15,17,15,19,21,17,18,22,17,13],'ReportConfirmations','Source: Service Confirmations.XLSX / Confirmations. Original status and Retail sales-order filters retained.');
format(report,rn,['F','H'],'dd-mmm-yyyy');format(report,rn,['G'],'mmm yyyy');format(report,rn,['I','J'],num);format(report,rn,['A','B','C','E','M'],'@');
const inum=table(inv,d.invoice,['Invoice Number','Item','Billing Date','Billing Month','Sales Order','Matched TicketID','Plant','Billing Type','Invoice ex GST','Invoice incl GST','GL Account','Material','Source Row'],[21,10,15,17,17,22,12,16,21,21,17,20,13],'SourceInvoices','Source: Service Confirmations.XLSX / ZFIR025. Includes negative invoice reversals.');
format(inv,inum,['C'],'dd-mmm-yyyy');format(inv,inum,['D'],'mmm yyyy');format(inv,inum,['I','J'],num);format(inv,inum,['A','B','E','F','G','K'],'@');
const dr=c=>`'Dashboard Tickets'!$${c}$5:$${c}$${dn}`,rr=c=>`'Report Details'!$${c}$5:$${c}$${rn}`;
const ir=c=>`'Invoice Lines'!$${c}$5:$${c}$${inum}`;
const mn=table(match,d.matched.map(x=>[x.ticket,null,x.center,x.order,x.so,null,null,null,date(x.month),null,null,null,null,null,null,null,null,null,null]),['TicketID','Location','Work Center','Order','AR SO','Invoice Number','Created Month','Billing Month','TECO Month','Dashboard Hours','Report Hours','Hours Difference','Dashboard Invoice','Report Invoice','Invoice Difference','Created Month Differs','Billing Month Differs','Source Invoice Net','Finding'],[13,18,17,16,17,21,17,17,17,22,20,22,23,22,23,24,24,23,63],'MatchedTickets','Matched by invoice number, sales order and service order. Differences are Dashboard minus Report.');
for(let i=0;i<d.matched.length;i++){
 const r=i+5,x=d.matched[i],s=x.row;
 for(const [to,from] of [['B','B'],['F','G'],['G','F'],['J','J'],['M','K']])match.getRange(`${to}${r}`).formulas=[[`='Dashboard Tickets'!${from}${s}`]];
 // Use actual Billing Date rather than the invoice inclusion month for timing comparisons.
 match.getRange(`H${r}`).formulas=[[`=DATE(YEAR('Dashboard Tickets'!H${s}),MONTH('Dashboard Tickets'!H${s}),1)`]];
 match.getRange(`K${r}:L${r}`).formulas=[[`=SUMIFS(${rr('I')},${rr('A')},$D${r},${rr('K')},"Yes")`,`=J${r}-K${r}`]];
 match.getRange(`N${r}:R${r}`).formulas=[[
 `=SUMIFS(${rr('J')},${rr('A')},$D${r},${rr('K')},"Yes")`,`=M${r}-N${r}`,
 `=IF(G${r}=I${r},"No","Yes")`,`=IF(H${r}=I${r},"No","Yes")`,`=SUMIFS(${ir('I')},${ir('E')},$E${r})`]];
 match.getRange(`S${r}`).values=[[x.ticket==='36147'?'Invoice 90043439 reversed by 90047125 on 27 Aug 2026.':null]];
}
format(match,mn,['G','H','I'],'mmm yyyy');format(match,mn,['J','K','L','M','N','O','R'],num);format(match,mn,['A','C','D','E','F'],'@');
const mr=c=>`'Matched Tickets'!$${c}$5:$${c}$${mn}`;
style(sum,55,'I');sum.getRange('A1').values=[['Dashboard vs Service Confirmations']];sum.getRange('A1').format.font={name:'Arial',size:13,bold:true,color:'#000000'};
sum.getRange('A2').values=[['July–September 2026. Labour in hours; invoices in AUD, excluding GST.']];
sum.getRange('A3').values=[['Monthly labour differences are large. Hours agree on all 76 matched tickets.']];
for(const [r,title] of [[4,'All dashboard locations'],[12,'Four common locations'],[21,'Matched tickets by completion month']]){sum.getRange(`A${r}`).values=[[title]];sum.getRange(`A${r}`).format.font.bold=true;}
sum.getRange('A13').values=[['Perth, Traralgon, Launceston and Geelong. Frankston and Other excluded.']];
for(const [h,start,isCommon] of [[6,7,false],[15,16,true]]){
 sum.getRange(`B${h-1}`).values=[['Labour Hours']];sum.getRange(`F${h-1}`).values=[['Invoice Amount']];sum.getRange(`A${h}:I${h}`).values=[['Month','Dashboard','Report','Difference','Difference %','Dashboard','Report','Difference','Difference %']];heading(sum,h,'I');
 for(let i=0;i<3;i++){
  const r=start+i,scope=isCommon?`,${dr('N')},"Yes"`:'';sum.getRange(`A${r}`).values=[[date(d.months[i])]];
  sum.getRange(`B${r}:I${r}`).formulas=[[
   `=SUMIFS(${dr('J')},${dr('F')},$A${r}${scope})`,`=SUMIFS(${rr('I')},${rr('G')},$A${r},${rr('K')},"Yes")`,`=B${r}-C${r}`,`=IF(C${r}=0,"n.a.",D${r}/C${r})`,
   `=SUMIFS(${dr('M')},${dr('I')},$A${r}${scope})`,`=SUMIFS(${rr('J')},${rr('G')},$A${r},${rr('L')},"Yes")`,`=F${r}-G${r}`,`=IF(G${r}=0,"n.a.",H${r}/G${r})`]];
 }
 const t=start+3;sum.getRange(`A${t}`).values=[['Total']];for(const c of ['B','C','D','F','G','H'])sum.getRange(`${c}${t}`).formulas=[[`=SUM(${c}${start}:${c}${t-1})`]];
 sum.getRange(`E${t}`).formulas=[[`=D${t}/C${t}`]];sum.getRange(`I${t}`).formulas=[[`=H${t}/G${t}`]];total(sum,t,'I');
 sum.getRange(`A${start}:A${t-1}`).setNumberFormat('mmm yyyy');sum.getRange(`B${start}:I${t}`).setNumberFormat(num);for(const c of ['E','I'])sum.getRange(`${c}${start}:${c}${t}`).setNumberFormat(pct);
}
sum.getRange('A22:H22').values=[['TECO Month','Tickets','Dashboard Hours','Report Hours','Difference','Dashboard Invoice','Report Invoice','Difference']];heading(sum,22,'H');
for(let i=0;i<2;i++){
 const r=i+23;sum.getRange(`A${r}`).values=[[date(d.months[i])]];
 sum.getRange(`B${r}:H${r}`).formulas=[[
 `=COUNTIFS(${mr('I')},A${r})`,`=SUMIFS(${mr('J')},${mr('I')},A${r})`,`=SUMIFS(${mr('K')},${mr('I')},A${r})`,`=C${r}-D${r}`,
 `=SUMIFS(${mr('M')},${mr('I')},A${r})`,`=SUMIFS(${mr('N')},${mr('I')},A${r})`,`=F${r}-G${r}`]];
}
sum.getRange('A25').values=[['Total']];sum.getRange('B25:H25').formulas=[Array.from({length:7},(_,i)=>`=SUM(${col(i+1)}23:${col(i+1)}24)`)];total(sum,25,'H');sum.getRange('A23:A24').setNumberFormat('mmm yyyy');sum.getRange('C23:H25').setNumberFormat(num);
const notes=[
 [27,'48 of the 76 matched tickets fall in different creation and completion months; 5 have different billing and completion months.'],
 [28,'Invoice amounts agree on 75 tickets. Ticket 36147 differs by $1,365.50: invoice reversed on 27 Aug 2026.'],
 [29,'Matched: 76 of 452 report orders. The remaining 376 orders have not been reconciled to dashboard tickets.'],
 [30,'Dashboard invoice data ends 20 Aug; report TECO dates end 9 Sep. September dashboard invoices are incomplete.'],
 [31,'Report totals: All Service labour and Retail Service invoices. Dashboard totals: All Ticket Types and All Invoices.'],
 [32,'Difference = Dashboard minus Report. Difference % uses Report as the base.'],
 [33,`Sources: dashboard snapshot ${d.snapshot}; Service Confirmations.XLSX.`],
 [34,'Future CreatedOn dates are retained from the dashboard source.']
];for(const [r,v] of notes)sum.getRange(`A${r}`).values=[[v]];
// A familiar location-by-month view to trace the same-site comparison.
sum.getRange('A36').values=[['Common locations by month']];sum.getRange('A36').format.font.bold=true;
sum.getRange('A37:H37').values=[['Location','Month','Dashboard Hours','Report Hours','Difference','Dashboard Invoice','Report Invoice','Difference']];heading(sum,37,'H');
let row=38;
for(const location of Object.values(d.centres))for(const m of d.months){
 const r=row++;sum.getRange(`A${r}:B${r}`).values=[[location,date(m)]];
 sum.getRange(`C${r}:H${r}`).formulas=[[
 `=SUMIFS(${dr('J')},${dr('B')},$A${r},${dr('F')},$B${r})`,`=SUMIFS(${rr('I')},${rr('D')},$A${r},${rr('G')},$B${r},${rr('K')},"Yes")`,`=C${r}-D${r}`,
 `=SUMIFS(${dr('M')},${dr('B')},$A${r},${dr('I')},$B${r})`,`=SUMIFS(${rr('J')},${rr('D')},$A${r},${rr('G')},$B${r},${rr('L')},"Yes")`,`=F${r}-G${r}`]];
}
sum.getRange('B38:B49').setNumberFormat('mmm yyyy');sum.getRange('C38:H49').setNumberFormat(num);
[19,19,22,22,22,24,24,22,22].forEach((w,i)=>sum.getRange(`${col(i)}1:${col(i)}55`).format.columnWidth=w);
sum.freezePanes.freezeRows(6);
wb.recalculate();
for(const [start,expected] of [[7,d.expected.all],[16,d.expected.common]]){
 const actual=sum.getRange(`B${start}:I${start+2}`).values;
 for(let i=0;i<3;i++)for(let j=0;j<8;j++)if(Math.abs(actual[i][j]-expected[i][j+1])>1e-7)throw Error(`Summary mismatch ${start+i},${j}`);
}
for(let i=0;i<d.matched.length;i++){
 const r=i+5,x=d.matched[i],v=match.getRange(`J${r}:R${r}`).values[0];
 if(Math.abs(v[1]-x.reportHours)>1e-7||Math.abs(v[4]-x.reportInvoice)>1e-7||Math.abs(v[2])>1e-7||Math.abs(v[8]-v[4])>1e-7)throw Error(`Matched ticket error ${x.ticket}`);
}
console.log((await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!',options:{useRegex:true,maxResults:10},maxChars:1000})).ndjson);
console.log((await wb.inspect({kind:'table',range:'Summary!A23:H25',tableMaxRows:3,tableMaxCols:8,maxChars:1400})).ndjson);
const outfile=path.join(dir,'Dashboard_vs_Service_Confirmations.xlsx');
await (await SpreadsheetFile.exportXlsx(wb)).save(outfile);
for(const [s,range,n] of [[sum,'A1:I19','comparison_summary'],[sum,'A21:I33','comparison_findings'],[match,'A1:L11','comparison_matched'],[match,'M1:S11','comparison_matched_amount'],[dash,'A1:M10','comparison_dashboard'],[report,'A1:O10','comparison_report'],[inv,'A1:M10','comparison_invoices']]){
 const img=await wb.render({sheetName:s.name,range,scale:1.5,format:'png'});await fs.writeFile(path.join(dir,n+'.png'),new Uint8Array(await img.arrayBuffer()));
}
console.log('SAVED',outfile);
