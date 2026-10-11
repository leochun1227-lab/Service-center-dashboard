"""Read-only SAP cancellation audit. Does not write dashboard inputs or code."""
from pathlib import Path
from decimal import Decimal
import collections
import json
import sys
sys.dont_write_bytecode = True
ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import pandas as pd
import pyodbc
import sap_invoice_enrich_from_hana as sap

OUT = ROOT / 'outputs' / 'invoice-validity-audit'
OUT.mkdir(exist_ok=True)
norm = sap.normalize_invoice_number
conn = pyodbc.connect(sap.load_hana_dsn(), timeout=60, autocommit=True)
queries = []
def query(sql, params=()):
    queries.append({'sql': sql, 'parameters': list(params)})
    cur = conn.cursor()
    cur.execute(sql, *params)
    cols = [c[0] for c in cur.description]
    return [{k: str(v).strip() if v is not None else '' for k, v in zip(cols, row)} for row in cur.fetchall()]

lookup = pd.read_excel(ROOT / 'c4c_ticket_table_z007_z010_checked_hana_final.xlsx', sheet_name='SAPInvoiceLookup', dtype=str).fillna('')
baseline = set(lookup.ERPInvoiceNumber.map(norm))
source = pd.read_excel('C:/Users/Leo.Li/Downloads/Service Department - SR URV Sept26.XLSX', sheet_name='Sheet1', dtype=str).fillna('')
materials = next(c for c in source.columns if c.strip().lower() == 'material')
doc_col = next(c for c in source.columns if c.strip().lower() == 'billing document')
file_ids = set(source.loc[source[materials].isin(['SOWTY999','SOWTY777']), doc_col].map(norm)) - {''}
service = query(f'''SELECT DISTINCT k."VBELN" FROM "{sap.SCHEMA}"."VBRK" k
 JOIN "{sap.SCHEMA}"."VBRP" p ON p."MANDT"=k."MANDT" AND p."VBELN"=k."VBELN"
 WHERE k."MANDT"=? AND p."MATNR" IN ('SOWTY999','SOWTY777')''', (sap.SAP_CLIENT,))
live_ids = {r['VBELN'] for r in service}
universe = baseline | live_ids | file_ids
headers = {}
for batch in sap.chunked(sorted(universe), 200):
    marks = ','.join('?' for _ in batch)
    rows = query(f'''SELECT "VBELN","FKART","FKDAT","NETWR","WAERK","FKSTO","SFAKN","VBTYP","ERDAT","ERZET"
      FROM "{sap.SCHEMA}"."VBRK" WHERE "MANDT"=? AND ("VBELN" IN ({marks}) OR "SFAKN" IN ({marks}))''', (sap.SAP_CLIENT,*batch,*batch))
    headers.update({r['VBELN']: r for r in rows})
missing = {r['SFAKN'] for r in headers.values() if r['SFAKN']} - headers.keys()
for batch in sap.chunked(sorted(missing),200):
    rows = query(f'''SELECT "VBELN","FKART","FKDAT","NETWR","WAERK","FKSTO","SFAKN","VBTYP","ERDAT","ERZET"
      FROM "{sap.SCHEMA}"."VBRK" WHERE "MANDT"=? AND "VBELN" IN ({','.join('?' for _ in batch)})''',(sap.SAP_CLIENT,*batch))
    headers.update({r['VBELN']:r for r in rows})
items=[]
for batch in sap.chunked(sorted(headers),200):
    items.extend(query(f'''SELECT "VBELN","POSNR","AUBEL","AUPOS","VGBEL","VGPOS","MATNR","NETWR"
      FROM "{sap.SCHEMA}"."VBRP" WHERE "MANDT"=? AND "VBELN" IN ({','.join('?' for _ in batch)})''',(sap.SAP_CLIENT,*batch)))
conn.close()
cancel_by_original=collections.defaultdict(list)
for r in headers.values():
    if r['SFAKN']:
        cancel_by_original[r['SFAKN']].append(r)
orders=collections.defaultdict(set)
for r in items:
    if r['AUBEL']:
        orders[r['VBELN']].add(r['AUBEL'])
def amount(r): return Decimal(r['NETWR'])
def cancelled(r): return r['FKART'] in {'S1','AB'} or r['VBTYP']=='N' or bool(r['SFAKN'])
def valid(r): return not cancelled(r) and r['FKSTO']!='X' and not cancel_by_original.get(r['VBELN'])
def signed(r): return -amount(r) if cancelled(r) else amount(r)
def stats(ids,prefix):
    rows=[headers[k] for k in ids if k in headers and headers[k]['FKDAT'].startswith(prefix)]
    active=[r for r in rows if valid(r)]
    by_order=collections.defaultdict(list)
    for r in active:
        order_key=tuple(sorted(orders[r['VBELN']])) or ('UNLINKED-'+r['VBELN'],)
        by_order[order_key].append(r)
    latest=[max(group,key=lambda r:(r['FKDAT'],r['ERDAT'],r['ERZET'],r['VBELN'])) for group in by_order.values()]
    pairs=[(headers.get(r['SFAKN']),r) for r in rows if cancelled(r)]
    return {
      'documents':len(rows),'signed_total':str(sum(map(signed,rows),Decimal(0))),
      'valid_count':len(active),'valid_total':str(sum(map(amount,active),Decimal(0))),
      'cancel_count':sum(cancelled(r) for r in rows),
      'cancelled_original_count':sum(not cancelled(r) and not valid(r) for r in rows),
      'latest_per_order_count':len(latest),'latest_per_order_total':str(sum(map(amount,latest),Decimal(0))),
      'multiple_valid_orders':[{ 'orders':list(k),'docs':[(r['VBELN'],r['FKDAT'],r['NETWR']) for r in group]} for k,group in by_order.items() if len(group)>1],
      'cancel_missing_original':[c['VBELN'] for o,c in pairs if o is None],
      'cancel_value_mismatch':[(o['VBELN'],c['VBELN']) for o,c in pairs if o and amount(o)!=amount(c)],
      'incoming_cross_period':[{'original':o,'cancel':c} for o,c in pairs if o and not o['FKDAT'].startswith(prefix)],
      'outgoing_cross_period':[{'original':r,'cancel':c} for r in rows if not cancelled(r) for c in cancel_by_original.get(r['VBELN'],[]) if not c['FKDAT'].startswith(prefix)],
      'valid_without_order':[r['VBELN'] for r in active if not orders[r['VBELN']]],
      'multi_order_invoices':[r['VBELN'] for r in active if len(orders[r['VBELN']])>1],
      'valid_ids':[r['VBELN'] for r in active],
    }
report={
 'file_september':stats(file_ids,'202609'),
 'saved_september':stats(baseline,'202609'),
 'saved_2026':stats(baseline,'2026'),
 'live_september':stats(live_ids,'202609'),
 'live_2026':stats(live_ids,'2026'),
 'saved_months':{month:stats(baseline,month) for month in sorted({r['FKDAT'][:6] for k,r in headers.items() if k in baseline and r['FKDAT'].startswith('2026')})},
 'lookup_missing':sorted(baseline-headers.keys()),
 'currencies':dict(collections.Counter(r['WAERK'] for r in headers.values())),
 'query_time':pd.Timestamp.now(tz='Australia/Sydney').isoformat(),
}
(OUT/'results.json').write_text(json.dumps(report,indent=2),encoding='utf-8')
(OUT/'queries.json').write_text(json.dumps(queries,indent=2),encoding='utf-8')
(OUT/'sap_headers.json').write_text(json.dumps(list(headers.values()),indent=2),encoding='utf-8')
(OUT/'sap_items.json').write_text(json.dumps(items,indent=2),encoding='utf-8')
for key in ['file_september','saved_september','saved_2026','live_september','live_2026']:
    print(key,json.dumps({k:v for k,v in report[key].items() if k!='valid_ids'}))
print('missing',report['lookup_missing'],'currencies',report['currencies'])
