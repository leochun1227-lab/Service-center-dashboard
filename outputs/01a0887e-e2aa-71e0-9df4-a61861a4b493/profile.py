import openpyxl,json,collections,datetime
from pathlib import Path
base=Path(__file__).resolve().parent
src=Path('c4c_ticket_table_z007_z010_checked_hana_final.xlsx')
w=openpyxl.load_workbook(src,read_only=True,data_only=True)
ts=[]
for sn in ['Tickets','NotAssigned']:
 s=w[sn]; rows=list(s.values); h=rows[0]
 ts.extend(dict(zip(h,r),source_sheet=sn,source_row=i) for i,r in enumerate(rows[1:],2) if r[0])
def norm(x): return str(x or '').strip().lstrip('0')
def num(x): return float(str(x or 0).replace(',',''))
print('TICKETS',len(ts),'unique',len(set(t['TicketID'] for t in ts)))
for k in ['DealerName','TicketStatusText']:
 print(k,collections.Counter(t[k] for t in ts))
for k in ['ERPInvoiceNumber','ERPFreeOrder','Billing date']:
 vals=[t[k] for t in ts if t[k]]
 print(k,len(vals),'duplicates',[(a,b) for a,b in collections.Counter(vals).items() if b>1][:20])
print('Dated sample',[[t.get(k) for k in ['TicketID','DealerName','ERPInvoiceNumber','ERPInvoiceNumberPrice','Billing date','TotalLabourHours','ERPFreeOrder']] for t in ts if t['ERPInvoiceNumber']][:10])
a=openpyxl.load_workbook('C:/Users/Leo.Li/Downloads/Service Confirmations.XLSX',read_only=True,data_only=True)
cr=list(a['Confirmations'].values);zr=list(a['ZFIR025'].values)
cs=[dict(zip(cr[0],r),source_row=i) for i,r in enumerate(cr[1:],2) if r[0]]
zs=[dict(zip(zr[0],r),source_row=i,BillingItem=r[16]) for i,r in enumerate(zr[1:],2) if r[0]]
print('CONFIRMATIONS',len(cs),'hours',sum(num(r['Activity to conf. 1 (ILE01)']) for r in cs),'orders',len(set(r['Order'] for r in cs)))
for k in ['Status','Work Center']:
 print(k,collections.Counter(r[k] for r in cs))
for date in ['TECO Date','Posting Date']:
 agg=collections.defaultdict(lambda:[0,0,0])
 for r in cs:
  d=r[date]; m=d.strftime('%Y-%m') if isinstance(d,datetime.datetime) else str(d)
  agg[m][0]+=num(r['Activity to conf. 1 (ILE01)']);agg[m][1]+=num(r['Confirmed Yield (GMEIN)']);agg[m][2]+=num(r['Invoice Amount'])
 print(date,dict(agg))
ivset={norm(t['ERPInvoiceNumber']) for t in ts if t['ERPInvoiceNumber']}
soset={norm(t['ERPFreeOrder']) for t in ts if t['ERPFreeOrder']}
print('MATCH invoices',sum(norm(z['Billing Document']) in ivset for z in zs),'/',len(zs),'CONF SO',sum(norm(c['AR SO']) in soset for c in cs))
print('Billing range',min(str(t['Billing date']) for t in ts if t['Billing date']),max(str(t['Billing date']) for t in ts if t['Billing date']))
print('Total labour',sum(num(t['TotalLabourHours']) for t in ts),'invoiced',sum(num(t['TotalLabourHours']) for t in ts if t['ERPInvoiceNumber']),'amount',sum(num(t['ERPInvoiceNumberPrice']) for t in ts))
print('ZFIR dup key',[(k,n) for k,n in collections.Counter((z['Billing Document'],z['Item']) for z in zs).items() if n>1][:10])
base.joinpath('sources.json').write_text(json.dumps({'tickets':ts,'confirmations':cs,'billing':zs},default=lambda x:x.isoformat() if isinstance(x,datetime.datetime) else str(x),ensure_ascii=False),encoding='utf-8')
