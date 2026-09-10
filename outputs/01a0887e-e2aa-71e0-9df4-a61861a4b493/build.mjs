import fs from 'node:fs/promises';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {Workbook,SpreadsheetFile} from '@oai/artifact-tool';
const dir=path.dirname(fileURLToPath(import.meta.url));
const d=JSON.parse(await fs.readFile(path.join(dir,'data.json'),'utf8'));
const labels=JSON.parse(await fs.readFile(path.join(dir,'english_labels.json'),'utf8'));
for(const rows of [d.tickets,d.confirmations,d.billing,d.expected])for(const row of rows)for(let i=0;i<row.length;i++)if(typeof row[i]==='string' && labels[row[i]])row[i]=labels[row[i]];
const wb=Workbook.create();
const summary=wb.worksheets.add('Monthly Summary');
const tickets=wb.worksheets.add('Ticket Details');
const conf=wb.worksheets.add('Confirmation Details');
const billing=wb.worksheets.add('Invoice Details');
const col=n=>{let x='';for(n++;n;n=Math.floor((n-1)/26))x=String.fromCharCode(65+(n-1)%26)+x;return x};
const dt=x=>typeof x==='string'&&/^\d{4}-\d{2}-\d{2}$/.test(x)?new Date(x+'T00:00:00Z'):x;
const money='#,##0.00;[Black](#,##0.00);0.00';
function base(sh,range){sh.showGridLines=false;sh.getRange(range).format.font={name:'Arial',size:10,color:'#000000'};sh.getRange(range).format.rowHeight=20;sh.getRange(range).format.verticalAlignment='center';}
function header(sh,r,last){const q=sh.getRange(`A${r}:${last}${r}`);q.format.font.bold=true;q.format.wrapText=true;q.format.horizontalAlignment='center';q.format.rowHeight=36;q.format.borders={bottom:{style:'thin',color:'#000000'}};}
function title(sh,r,text){sh.getRange(`A${r}`).values=[[text]];sh.getRange(`A${r}`).format.font={name:'Arial',size:13,bold:true,color:'#000000'};sh.getRange(`A${r}`).format.rowHeight=24;}
function detail(sh,rows,heads,widths,datecols,numcols,source,name){
 const last=col(heads.length-1),end=rows.length+5;
 base(sh,`A1:${last}${end}`);title(sh,2,sh.name);
 sh.getRange('A3').values=[[source]];
 sh.getRange(`A5:${last}5`).values=[heads];
 sh.getRange(`A6:${last}${end}`).values=rows.map(r=>r.map(dt));
 for(let i=0;i<widths.length;i++)sh.getRange(`${col(i)}1:${col(i)}${end}`).format.columnWidth=widths[i];
 for(const c of datecols){sh.getRange(`${c}6:${c}${end}`).setNumberFormat(c==='E'||c==='H'&&sh===tickets?'yyyy-mm-dd':'yyyy-mm');sh.getRange(`${c}6:${c}${end}`).format.horizontalAlignment='center';}
 for(const c of numcols)sh.getRange(`${c}6:${c}${end}`).setNumberFormat(money);
 const tab=sh.tables.add(`A5:${last}${end}`,true,name);tab.style='TableStyleLight1';tab.showBandedColumns=false;
 sh.getRange(`A5:${last}${end}`).format.fill='#FFFFFF';
 header(sh,5,last);sh.freezePanes.freezeRows(5);sh.freezePanes.freezeColumns(2);
 return end;
}
const nt=detail(tickets,d.tickets,['TicketID','Dealer','Ticket Type','Ticket Status','CreatedOn','Created Month','ERPInvoiceNumber','Billing date','Invoice Month','Invoice ex GST (AUD)','TotalLabourHours (h)','TECO Month','TECO Date(s)','AR SO','Order','TECO Match Result','Source Sheet','Source Row'],[13,33,18,23,14,13,23,14,13,21,22,15,26,23,23,30,17,10],['E','F','H','I','L'],['J','K'],'Source: c4c_ticket_table_z007_z010_checked_hana_final.xlsx (Tickets + NotAssigned)','TicketDetails');
const nc=detail(conf,d.confirmations,['Order','Confirmation','Work Center','AR SO','TECO Date','Completion Month','Posting Date','Posting Month','Confirmed labour (h)','Invoice Amount (AUD)','Status','All Service Filter','Retail Service Filter','TECO Order Count','Posting Order Count','Retail Order Count','Billing Document(s)','Linked TicketID(s)','Source Row'],[14,16,16,14,14,13,15,13,23,23,48,19,20,19,19,19,25,24,10],['E','F','G','H'],['I','J'],'Source: Service Confirmations.XLSX / Confirmations. Negative labour reversals and original Invoice Amount values retained.','ConfirmationDetails');
conf.getRange(`G6:G${nc}`).setNumberFormat('yyyy-mm-dd');
const nb=detail(billing,d.billing,['Billing Document','Item','Billing Date','Invoice Month','Sales Order','Linked TicketID','Plant','Billing Type','Invoice ex GST (AUD)','Invoice incl GST (AUD)','GL Account','Material','Source Row'],[23,10,15,13,17,17,12,15,23,23,17,21,10],['C','D'],['I','J'],'Source: Service Confirmations.XLSX / ZFIR025. One row per invoice line item.','InvoiceDetails');
billing.getRange(`C6:C${nb}`).setNumberFormat('yyyy-mm-dd');
for(const [sh,cols,end] of [[tickets,['A','G','M','N','O'],nt],[conf,['A','B','C','D','Q','R'],nc],[billing,['A','B','E','F','G','K'],nb]])for(const c of cols)sh.getRange(`${c}6:${c}${end}`).setNumberFormat('@');
for(let r=6;r<=nc;r++){
 conf.getRange(`N${r}:P${r}`).formulas=[[
  `=IF(L${r}="Yes",IF(COUNTIFS($A$6:A${r},A${r},$C$6:C${r},C${r},$F$6:F${r},F${r},$L$6:L${r},"Yes")=1,1,0),0)`,
  `=IF(L${r}="Yes",IF(COUNTIFS($A$6:A${r},A${r},$C$6:C${r},C${r},$H$6:H${r},H${r},$L$6:L${r},"Yes")=1,1,0),0)`,
  `=IF(M${r}="Yes",IF(COUNTIFS($A$6:A${r},A${r},$C$6:C${r},C${r},$F$6:F${r},F${r},$M$6:M${r},"Yes")=1,1,0),0)`]];
}
const tr=c=>`'Ticket Details'!$${c}$6:$${c}$${nt}`;
const cr=c=>`'Confirmation Details'!$${c}$6:$${c}$${nc}`;
const br=c=>`'Invoice Details'!$${c}$6:$${c}$${nb}`;
base(summary,'A1:K85');
title(summary,2,'Monthly Ticket Labour and Invoice Summary');
summary.getRange('A3').values=[['Labour in hours. Amounts in AUD, excluding GST. All ticket statuses included; missing dates shown separately.']];
summary.getRange('B4').values=[['Created month (CreatedOn)']];summary.getRange('D4').values=[['Invoice month (Billing date)']];summary.getRange('G4').values=[['Completion month (TECO Date, matched tickets)']];
summary.getRange('A4:J4').format.font.bold=true;
summary.getRange('A5:I5').values=[['Month','Ticket Count','Total labour (h)','Ticket Count','Total labour (h)','Invoice (AUD)','Ticket Count','Total labour (h)','Invoice (AUD)']];
header(summary,5,'I');
const keys=[...d.months,'No Date','Unmatched','To Be Confirmed'];
summary.getRange(`A6:A${keys.length+5}`).values=keys.map(k=>[dt(k)]);summary.getRange(`A6:A${keys.length+5}`).setNumberFormat('yyyy-mm');
for(let i=0;i<keys.length;i++){
 const r=i+6;
 summary.getRange(`B${r}:I${r}`).formulas=[[
  `=COUNTIFS(${tr('F')},$A${r})`,`=SUMIFS(${tr('K')},${tr('F')},$A${r})`,
  `=COUNTIFS(${tr('I')},$A${r})`,`=SUMIFS(${tr('K')},${tr('I')},$A${r})`,`=SUMIFS(${tr('J')},${tr('I')},$A${r})`,
  `=COUNTIFS(${tr('L')},$A${r})`,`=SUMIFS(${tr('K')},${tr('L')},$A${r})`,`=SUMIFS(${tr('J')},${tr('L')},$A${r})`]];
}
const total=keys.length+6;
summary.getRange(`A${total}`).values=[['Total']];summary.getRange(`B${total}:I${total}`).formulas=[Array.from({length:8},(_,i)=>`=SUM(${col(i+1)}6:${col(i+1)}${total-1})`)];
summary.getRange(`A${total}:I${total}`).format.font.bold=true;summary.getRange(`A${total}:I${total}`).format.borders={top:{style:'thin',color:'#000000'}};
summary.getRange(`A${total+2}`).values=[['TECO linked via invoice number, AR SO and Order. 76 tickets have a completion month; the remainder are unmatched.']];
summary.getRange(`A${total+3}`).values=[['Ticket billing dates end on 2026-08-20. Future CreatedOn dates are retained as provided.']];
// Exact original pivot recreation, preserving its selected status and sales-order filters.
const section=total+6,hr=section+2,first=hr+1;
title(summary,section,'Source Summary Reconciliation by TECO Date and Work Center');
summary.getRange(`A${section+1}`).values=[['Original TECO status and Retail AR SO filters retained. Orders counted once per group.']];
summary.getRange(`A${hr}:K${hr}`).values=[['Scope','Month','Work Center','Recalculated Labour (h)','Original Labour (h)','Labour Difference (h)','Recalculated Orders','Original Orders','Recalculated Invoice','Original Invoice','Invoice Difference']];header(summary,hr,'K');
for(let i=0;i<d.original.length;i++){
 const r=first+i,o=d.original[i],flag=o[0]==='All Service'?'L':'M',counter=o[0]==='All Service'?'N':'P';
 summary.getRange(`A${r}:K${r}`).values=[[o[0],dt(o[1]),o[2],null,o[3],null,null,o[4],null,o[5],null]];
 const criteria=`${cr('F')},$B${r},${cr('C')},$C${r},${cr(flag)},"Yes"`;
 summary.getRange(`D${r}`).formulas=[[`=SUMIFS(${cr('I')},${criteria})`]];
 summary.getRange(`F${r}`).formulas=[[`=D${r}-E${r}`]];
 summary.getRange(`G${r}`).formulas=[[`=SUMIFS(${cr(counter)},${criteria})`]];
 if(o[0]!=='All Service'){
  summary.getRange(`I${r}`).formulas=[[`=SUMIFS(${cr('J')},${criteria})`]];
  summary.getRange(`K${r}`).formulas=[[`=I${r}-J${r}`]];
 }
}
const last=first+d.original.length-1;
summary.getRange(`B${first}:B${last}`).setNumberFormat('yyyy-mm');summary.getRange(`D${first}:F${last}`).setNumberFormat(money);summary.getRange(`I${first}:K${last}`).setNumberFormat(money);
summary.getRange(`B${first}:B${last}`).format.horizontalAlignment='center';
const pr=last+4;
title(summary,pr,'Source Monthly Totals by Completion, Posting and Billing Date');
summary.getRange(`A${pr+1}`).values=[['Labour uses the original All Service status filter. ZFIR025 includes all invoice lines by Billing Date.']];
summary.getRange(`A${pr+2}:F${pr+2}`).values=[['Month','TECO Labour (h)','Posting Labour (h)','Labour Difference (h)','ZFIR025 Invoice ex GST','ZFIR025 Invoice incl GST']];header(summary,pr+2,'F');
const am=[...new Set(d.billing.map(r=>r[3]).concat(d.confirmations.map(r=>r[5])))].filter(m=>/^\d/.test(m)).sort();
for(let i=0;i<am.length;i++){
 const r=pr+3+i;summary.getRange(`A${r}`).values=[[dt(am[i])]];
 summary.getRange(`B${r}:F${r}`).formulas=[[
 `=SUMIFS(${cr('I')},${cr('F')},$A${r},${cr('L')},"Yes")`,
 `=SUMIFS(${cr('I')},${cr('H')},$A${r},${cr('L')},"Yes")`,`=B${r}-C${r}`,
 `=SUMIFS(${br('I')},${br('D')},$A${r})`,`=SUMIFS(${br('J')},${br('D')},$A${r})`]];
}
summary.getRange(`A${pr+3}:A${pr+2+am.length}`).setNumberFormat('yyyy-mm');summary.getRange(`B${pr+3}:F${pr+2+am.length}`).setNumberFormat(money);
for(let i=0;i<11;i++)summary.getRange(`${col(i)}1:${col(i)}85`).format.columnWidth=[25,19,21,21,24,24,20,21,24,24,22][i];
for(const c of ['C','E','F','H','I'])summary.getRange(`${c}6:${c}${total}`).setNumberFormat(money);
summary.freezePanes.freezeRows(5);
wb.recalculate();
const vals=summary.getRange(`A6:I${total-1}`).values;
for(let i=0;i<d.expected.length;i++)for(let j=1;j<9;j++)if(Math.abs(vals[i][j]-d.expected[i][j])>1e-6)throw Error(`Mismatch ${i},${j}: ${vals[i][j]} vs ${d.expected[i][j]}`);
for(let r=first;r<=last;r++){
 const v=summary.getRange(`D${r}:K${r}`).values[0];
 if(Math.abs(v[2])>1e-6||v[3]!==v[4]||(typeof v[7]==='number'&&Math.abs(v[7])>1e-6))throw Error(`Original summary mismatch ${r}`);
}
console.log((await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!',options:{useRegex:true,maxResults:10},summary:'Formula check',maxChars:1500})).ndjson);
console.log((await wb.inspect({kind:'table',range:`Monthly Summary!A${total-3}:I${total}`,include:'values,formulas',tableMaxRows:5,tableMaxCols:9,maxChars:1800})).ndjson);
await (await SpreadsheetFile.exportXlsx(wb)).save(path.join(dir,'Ticket_Labour_Invoice_Monthly_Reconciliation_EN.xlsx'));
for(const [sh,range,name] of [[summary,'A2:I13','summary'],[summary,`A${section}:K${first+5}`,'reconciliation'],[tickets,'A2:L12','tickets'],[conf,'A2:M12','confirmations'],[billing,'A2:M12','invoices']]){
 try{const img=await wb.render({sheetName:sh.name,range,scale:1.5,format:'png'});await fs.writeFile(path.join(dir,name+'_en.png'),new Uint8Array(await img.arrayBuffer()));}catch(e){console.log('RENDER ERROR',name,e.message);}
}
await fs.writeFile(path.join(dir,'validation_en.json'),JSON.stringify({ticketCount:d.tickets.length,confirmationRows:d.confirmations.length,billingRows:d.billing.length,originalMatchedRows:d.original.length,totalRow:total,originalFirst:first,originalLast:last,summary:vals},null,2));
console.log('DONE',path.join(dir,'Ticket_Labour_Invoice_Monthly_Reconciliation_EN.xlsx'));
