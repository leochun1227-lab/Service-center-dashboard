"""Read-only, bounded C4C probes. Saves aggregate results, never credentials or ticket bodies."""
import ast
import json
import os
import re
import sys
from datetime import datetime, timezone
from pathlib import Path

import requests

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

document = Path(r'C:\Users\Leo.Li\Downloads\BI工单查询接口交接文档.md').read_text(encoding='utf-8-sig')
def table_fields(section):
    block = document.split('## ' + section + '.', 1)[1].split('\n## ', 1)[0].split('\n### ', 1)[0]
    return re.findall(r'^\| ([A-Za-z][A-Za-z0-9]+) \|', block, re.M)

expected = {name: table_fields(section) for name, section in [('ticket', '6'), ('issue', '7'), ('parts', '8'), ('worklogs', '9')]}
def profile(rows, fields):
    return {field: {
        'missing': sum(field not in row for row in rows),
        'null': sum(field in row and row[field] is None for row in rows),
        'empty': sum(field in row and row[field] in ('', [], {}) for row in rows),
        'nonempty': sum(field in row and row[field] is not None and row[field] not in ('', [], {}) for row in rows),
    } for field in fields}

session = requests.Session()
session.trust_env = False
output = {'at_utc': datetime.now(timezone.utc).isoformat(), 'checks': []}
destination = Path(__file__).with_name('field_checks_filter.json')
for ticket_type in ['Z007', 'Z011']:
    params = {'$top': 5, '$skip': 0, '$filter': f"(CCSRQ_DPY_ROLE_CD eq '1001') and (CDOC_PROC_TY eq '{ticket_type}')"}
    result = {'ticket_type_requested': ticket_type, 'endpoint': settings['BASE_URL'].rstrip('/') + settings['PATH'], 'params': params}
    try:
        response = session.get(result['endpoint'], params=params, auth=(settings['USERNAME'], settings['PASSWORD']), timeout=(15, 60), allow_redirects=False)
        result.update(status=response.status_code, elapsed_seconds=round(response.elapsed.total_seconds(), 2))
        if response.status_code == 200:
            payload = response.json()
            rows = payload.get('data', [])
            result.update(count=payload.get('count'), returned=len(rows), actual_types=sorted({row.get('TicketType', '') for row in rows}), ticket_fields=profile(rows, expected['ticket']))
            issues = [issue for row in rows for issue in (row.get('IssueItems') or []) if isinstance(issue, dict)]
            parts = [part for issue in issues for part in (issue.get('PartsItems') or []) if isinstance(part, dict)]
            logs = [log for issue in issues for log in (issue.get('WorkLogs') or []) if isinstance(log, dict)]
            result.update(issue_count=len(issues), issue_fields=profile(issues, expected['issue']), parts_count=len(parts), worklogs_count=len(logs))
            result['all_present_ticket_fields'] = sorted({key for row in rows for key in row})
            result['classification_examples'] = [{key: issue.get(key) for key in ['IssuesPosition', 'IssuesPositionText', 'Subcategory', 'SubcategoryText', 'SubcategoryReason', 'SubcategoryReasonText', 'RepairItem', 'RepairItemText']} for issue in issues[:10]]
    except requests.RequestException as error:
        result['error_type'] = type(error).__name__
    except (ValueError, TypeError, AttributeError) as error:
        result['response_error_type'] = type(error).__name__
    output['checks'].append(result)
    destination.write_text(json.dumps(output, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps({key: value for key, value in result.items() if key not in ['ticket_fields', 'issue_fields']}, ensure_ascii=True), flush=True)
