import json,collections,datetime
from pathlib import Path
p=Path(__file__).resolve().parent
d=json.loads((p/'data.json').read_text(encoding='utf-8'))
raw=json.loads((p/'sources.json').read_text(encoding='utf-8'))
s=Path('dashboard-data.js').read_text(encoding='utf-8')
web=json.loads(s[s.index('{'):].strip().removesuffix(';'))
source={str(t['TicketID']):t for t in raw['tickets']}
detail={str(t['ticketId']):t for t in web['pages']['overview']['ticketDetails']}
dealer={'Regent RV - Perth':'Perth','Regent RV - Traralgon':'Traralgon','Snowy River Launceston':'Launceston','Snowy River Geelong':'Geelong','Regent RV - Frankston':'Frankston'}
centres={'32110001':'Perth','32110002':'Traralgon','32110003':'Launceston','32110004':'Geelong'}
rows=[]
for t in d['tickets']:
    src=source[t[0]];status=(t[3] or '').lower();inv=t[6];amount=t[9];billing=t[8]
    placeholder=status=='create invoice' and (not inv or amount is None)
    valid=billing!='无日期' and bool(inv) and amount is not None and status not in {'cancel','cancel invoice'}
    eligible=valid or placeholder
    invoice_month=(billing if billing!='无日期' else t[5]) if eligible else 'Excluded'
    included=(0 if placeholder else amount) if eligible else 0
    yard=dealer.get(t[1],'Other')
    match=t[11].startswith('2026')
    rows.append([t[0],yard,t[2],t[3],t[4],t[5],t[6],t[7],invoice_month,t[10],amount,'Yes' if eligible else 'No',included,'Yes' if yard in centres.values() else 'No',t[14] if match else None,t[11] if match else 'Unmatched',t[16],t[17]])
    assert abs(t[10]-detail[t[0]]['labourHours'])<1e-8
rows.sort(key=lambda r:(r[5],r[1],r[0]))
byticket={r[0]:i+5 for i,r in enumerate(rows)}
matched=[t for t in d['tickets'] if t[11].startswith('2026')]
byso={t[13]:t[0] for t in matched}
report=[]
for c in d['confirmations']:
    report.append([c[0],c[1],c[2],centres[c[2]],c[3],c[4],c[5] if c[5]!='无日期' else 'No Date',c[6],c[8],c[9],'Yes' if c[11]=='是' else 'No','Yes' if c[12]=='是' else 'No',byso.get(c[3]),'Matched' if c[3] in byso else 'Unmatched',c[18]])
report.sort(key=lambda r:(r[6],r[2],r[0],r[-1]))
mr=[]
for t in matched:
    cs=[c for c in report if c[4]==t[13] and c[10]=='Yes']
    assert cs and len({c[0] for c in cs})==1
    mr.append({'ticket':t[0],'row':byticket[t[0]],'order':t[14],'so':t[13],'center':cs[0][2],'month':t[11],'reportHours':sum(c[8] for c in cs),'reportInvoice':sum(c[9] or 0 for c in cs),'createdShift':t[5]!=t[11],'invoiceShift':t[8]!=t[11]})
mr.sort(key=lambda x:(x['month'],x['center'],x['ticket']))
months=['2026-07-01','2026-08-01','2026-09-01']
expect={}
for common in [False,True]:
    out=[]
    for m in months:
        chosen=[r for r in rows if not common or r[13]=='Yes']
        wh=sum(r[9] for r in chosen if r[5]==m)
        wi=sum(r[12] for r in chosen if r[8]==m)
        rh=sum(r[8] for r in report if r[6]==m and r[10]=='Yes')
        ri=sum(r[9] or 0 for r in report if r[6]==m and r[11]=='Yes')
        out.append([m,wh,rh,wh-rh,(wh-rh)/rh,wi,ri,wi-ri,(wi-ri)/ri])
        if not common:
            label=datetime.date.fromisoformat(m).strftime('%b %Y');page=web['pages']['overview']
            assert abs(wh-sum(v['totalHours'] for v in page['monthlyLabour'][label].values()))<1e-7
            assert abs(wi-sum(v['invoicedAmount'] for v in page['monthlyDealerActivityByType'][label]['All Ticket Types']))<1e-7
    expect['common' if common else 'all']=out
print('MONTHLY',expect)
print('MATCHED',len(mr),'hours',sum(r['reportHours'] for r in mr),'created shifts',sum(r['createdShift'] for r in mr),'invoice shifts',sum(r['invoiceShift'] for r in mr))
print('ORDER COVERAGE',len({r[0] for r in report if r[10]=='Yes'}),len({r[0] for r in report if r[11]=='Yes'}))
result={'dashboard':rows,'report':report,'matched':mr,'invoice':d['billing'],'months':months,'expected':expect,'centres':centres,'snapshot':web['meta']['lastUpdated']}
(p/'comparison_data.json').write_text(json.dumps(result,ensure_ascii=False),encoding='utf-8')
