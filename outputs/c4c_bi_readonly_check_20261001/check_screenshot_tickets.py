"""Read-only lookup of screenshot tickets, retaining only requested fields."""
import ast
import json
import os
import sys
from datetime import datetime, timezone
from pathlib import Path
import requests
from urllib.parse import quote

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))
from dashboard_credentials import apply_saved_settings
apply_saved_settings()
settings = {}
tree = ast.parse((ROOT / 'export_filtered_tickets_with_dealer_resolution.py').read_text(encoding='utf-8-sig'))
for node in tree.body:
    if isinstance(node, ast.Assign) and isinstance(node.value, ast.Call) and isinstance(node.value.func, ast.Attribute) and node.value.func.attr == 'getenv' and len(node.value.args) == 2:
        try:
            settings[node.targets[0].id] = os.getenv(ast.literal_eval(node.value.args[0]), ast.literal_eval(node.value.args[1]))
        except (ValueError, AttributeError):
            pass

session = requests.Session()
session.trust_env = False
out = {'at_utc': datetime.now(timezone.utc).isoformat(), 'checks': []}
for kind, target in [('Z007', '43301'), ('Z011', '43316')]:
    check = {'type': kind, 'target': target, 'pages': [], 'matches': []}
    for skip in [0, 1000]:
        params = {'$top': 1000, '$skip': skip, '$filter': f"(CCSRQ_DPY_ROLE_CD eq '1001') and (CDOC_PROC_TY eq '{kind}')"}
        try:
            response = session.get(settings['BASE_URL'].rstrip('/') + settings['PATH'] + f"?$top=1000&$skip={skip}&$filter=" + quote(params['$filter'], safe="()'"), auth=(settings['USERNAME'], settings['PASSWORD']), timeout=(15, 60), allow_redirects=False)
            page = {'params': params, 'status': response.status_code}
            check['pages'].append(page)
            if response.status_code != 200:
                break
            data = response.json()
            rows = data.get('data', [])
            page.update(page_size=data.get('pageSize'), page_number=data.get('pageNumber'), count=data.get('count'), returned=len(rows), issueitems_present=sum('IssueItems' in x for x in rows), subticket_present=sum('SubticketType' in x for x in rows), customer_role_present=sum(any(str(p.get('InvolvedPartyRoleID')) == '1001' and p.get('InvolvedPartyName') for p in (x.get('InvolvedParties') or [])) for x in rows))
            for row in rows:
                if str(row.get('TicketID', '')).lstrip('0') != target:
                    continue
                check['matches'].append({
                    'TicketID': row.get('TicketID'), 'TicketType': row.get('TicketType'), 'TicketTypeText': row.get('TicketTypeText'),
                    'customers': [p for p in (row.get('InvolvedParties') or []) if str(p.get('InvolvedPartyRoleID')) == '1001'],
                    'field_names': sorted(row),
                    'requested_fields': {k: row[k] for k in row if any(t in k.lower() for t in ['subticket', 'issue', 'subcategory'])},
                })
            print(json.dumps({'type': kind, 'skip': skip, **page, 'target_matches': check['matches']}, ensure_ascii=True), flush=True)
            if check['matches'] or not rows or skip + 1000 >= int(data.get('count') or 0):
                break
        except (requests.RequestException, ValueError, TypeError) as error:
            check['error_type'] = type(error).__name__
            break
    out['checks'].append(check)
    Path(__file__).with_name('screenshot_ticket_results.json').write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding='utf-8')
