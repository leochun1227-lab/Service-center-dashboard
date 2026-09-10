import json,datetime,collections,zipfile,xml.etree.ElementTree as ET,re,openpyxl
from pathlib import Path
base=Path(__file__).resolve().parent
d=json.loads((base/'sources.json').read_text(encoding='utf-8'))
ns={'x':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
z=zipfile.ZipFile('C:/Users/Leo.Li/Downloads/Service Confirmations.XLSX')
p=ET.fromstring(z.read('xl/pivotTables/pivotTable1.xml'))
members=[e.attrib['name'] for e in p.findall('.//x:member',ns)]
retail={s.split('.&[')[-1][:-1] for s in members if '[AR SO]' in s}
statuses={s.split('.&[')[-1][:-1] for s in members if '[Status]' in s}
def norm(x): return str(x or '').strip().lstrip('0')
def num(x): return None if x is None or x=='' else float(str(x).replace(',',''))
def date(x):
 if not x or str(x)=='00:00:00':return None
 for f in ['%Y-%m-%dT%H:%M:%S','%Y-%m-%d %H:%M:%S','%Y-%m-%d','%d/%m/%Y']:
  try:return datetime.datetime.strptime(str(x),f).date().isoformat()
  except ValueError:pass
 raise ValueError(x)
def month(x):return x[:7]+'-01' if x else '无日期'
byinv=collections.defaultdict(list);byso=collections.defaultdict(list)
for b in d['billing']:byinv[norm(b['Billing Document'])].append(b)
for c in d['confirmations']:byso[norm(c['AR SO'])].append(c)
ticketrows=[];timap={norm(t['ERPInvoiceNumber']):t['TicketID'] for t in d['tickets'] if t['ERPInvoiceNumber']}
for t in d['tickets']:
 inv=t['ERPInvoiceNumber'];bs=byinv.get(norm(inv),[]) if inv else []
 sos=sorted({norm(b['sales order']) for b in bs});cs=[c for so in sos for c in byso.get(so,[])]
 cdates=sorted({date(c['TECO Date']) for c in cs if date(c['TECO Date'])})
 cms={month(x) for x in cdates};orders=sorted({c['Order'] for c in cs})
 if not bs:tm='未匹配';state='附件无对应发票' if inv else '无发票号'
 elif not cs:tm='未匹配';state='附件无对应报工订单'
 elif not cdates:tm='无日期';state='附件无TECO日期'
 elif any(not date(c['TECO Date']) for c in cs):tm='待确认';state='部分订单无TECO日期'
 elif len(cms)>1:tm='待确认';state='关联多个完工月份'
 else:tm=next(iter(cms));state='发票→销售订单→报工订单'
 created=date(t['CreatedOn']);billing=date(t['Billing date'])
 ticketrows.append([t['TicketID'],t['DealerName'],t['TicketTypeText'],t['TicketStatusText'],created,month(created),inv,billing,month(billing),num(t['ERPInvoiceNumberPrice']),num(t['TotalLabourHours']),tm,', '.join(cdates),', '.join(sos),', '.join(orders),state,t['source_sheet'],t['source_row']])
ticketrows.sort(key=lambda r:(r[5],r[1] or '',str(r[0])))
cr=[]
seen_teco=set();seen_post=set();seen_retail=set()
for c in sorted(d['confirmations'],key=lambda c:(month(date(c['TECO Date'])),c['Work Center'],c['Order'],c['source_row'])):
 teco=date(c['TECO Date']);post=date(c['Posting Date']);inc=c['Status'] in statuses;ret=inc and c['AR SO'] in retail
 key=(month(teco),c['Work Center'],c['Order']);pk=(month(post),c['Work Center'],c['Order'])
 uniq=int(inc and key not in seen_teco);up=int(inc and pk not in seen_post);ur=int(ret and key not in seen_retail)
 if inc:seen_teco.add(key);seen_post.add(pk)
 if ret:seen_retail.add(key)
 invs=sorted({b['Billing Document'] for b in d['billing'] if norm(b['sales order'])==norm(c['AR SO'])})
 tids=sorted({timap[norm(v)] for v in invs if norm(v) in timap})
 cr.append([c['Order'],c['Confirmation'],c['Work Center'],c['AR SO'],teco,month(teco),post,month(post),num(c['Activity to conf. 1 (ILE01)']),num(c['Invoice Amount']),c['Status'],'是' if inc else '否','是' if ret else '否',uniq,up,ur,', '.join(invs),', '.join(tids),c['source_row']])
br=[]
for b in d['billing']:
 dt=date(b['Billing Date'])
 br.append([b['Billing Document'],b['BillingItem'],dt,month(dt),b['sales order'],timap.get(norm(b['Billing Document'])),b['Plant'],b['Billing Type'],num(b['Sales price(Excl.GST)- AUD']),num(b['ales price(Incl.GST)- AUD']),b['Account Number'],b['Material'],b['source_row']])
br.sort(key=lambda r:(r[3],r[0],r[1]))
allmonths=sorted({r[i] for r in ticketrows for i in [5,8,11] if re.match(r'^\d{4}',r[i])}|{r[5] for r in cr if re.match(r'^\d{4}',r[5])})
orig=[]
w=openpyxl.load_workbook('C:/Users/Leo.Li/Downloads/Service Confirmations.XLSX',data_only=True,read_only=True)
rows=list(w['Summary'].values)
for start,end,scope in [(4,20,'All Service'),(28,44,'Retail Service only')]:
 wc=None
 for idx in range(start,end):
  r=rows[idx]
  if str(r[0]).startswith('3211'):wc=r[0]
  elif r[0] in ['Jul','Aug','Sep']:orig.append([scope,'2026-'+{'Jul':'07','Aug':'08','Sep':'09'}[r[0]]+'-01',wc,r[1],r[2],r[3]])
for scope in ['All Service','Retail Service only']:
 for o in [x for x in orig if x[0]==scope]:
  subset=[r for r in cr if r[5]==o[1] and r[2]==o[2] and r[11 if scope=='All Service' else 12]=='是']
  hours=sum(r[8] or 0 for r in subset);cnt=len(set(r[0] for r in subset));amt=sum(r[9] or 0 for r in subset)
  assert abs(hours-o[3])<1e-7,(o,hours)
  assert cnt==o[4],(o,cnt)
  if scope!='All Service':assert abs(amt-o[5])<1e-7,(o,amt)
expected=[]
for m in allmonths+['无日期','未匹配','待确认']:
 a=[r for r in ticketrows if r[5]==m];b=[r for r in ticketrows if r[8]==m];c=[r for r in ticketrows if r[11]==m]
 expected.append([m,len(a),sum(r[10] or 0 for r in a),len(b),sum(r[10] or 0 for r in b),sum(r[9] or 0 for r in b),len(c),sum(r[10] or 0 for r in c),sum(r[9] or 0 for r in c)])
result={'tickets':ticketrows,'confirmations':cr,'billing':br,'months':allmonths,'original':orig,'expected':expected,'statuses':sorted(statuses),'retail_so_count':len(retail)}
(base/'data.json').write_text(json.dumps(result,ensure_ascii=False),encoding='utf-8')
print('Original summary reproduced exactly:',len(orig),'rows')
print('TECO mapping',collections.Counter(r[15] for r in ticketrows))
print('Monthly',expected[-7:])
print('Dates',min(r[4] for r in ticketrows if r[4]),max(r[4] for r in ticketrows if r[4]))
