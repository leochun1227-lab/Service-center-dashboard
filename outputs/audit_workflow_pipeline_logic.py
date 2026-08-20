from __future__ import annotations

import json
from collections import defaultdict
from pathlib import Path
from typing import Any

import pandas as pd


BASE = Path(r"C:\Users\Leo.Li\Documents\ChatGPT\Service centre dashboard")
SOURCE = BASE / "c4c_ticket_table_z007_z010_checked_hana_final.xlsx"
DATA_JS = BASE / "dashboard-data.js"
OUT = Path("outputs") / "workflow_pipeline_audit.json"

DEALER_LABELS = {
    "Regent RV - Perth": "Perth",
    "Regent RV - Traralgon": "Traralgon",
    "Snowy River Launceston": "Launceston",
    "Snowy River Geelong": "Geelong",
    "Regent RV - Frankston": "Frankston",
}
DEALERS = ["Perth", "Traralgon", "Launceston", "Geelong", "Frankston"]
TYPES = ["Repair ticket", "PDI ticket"]
REPAIR_STAGES = [
    "Awaiting Quote Approval",
    "Approved, Awaiting Repair",
    "Repair In Progress",
    "Repair Complete, Awaiting Time Claim",
    "Time Claimed, Awaiting Invoice",
    "Completed / Invoiced",
    "Invoice Cancelled, SO Still Open",
    "Cancelled",
]
PDI_STAGES = [
    "Awaiting PDI Start",
    "PDI Complete, Awaiting Time Claim",
    "PDI Completed",
    "Cancelled",
]
INTERNAL_PARTY_NAME = "REGENT RV PTY LTD"


def clean(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    return "" if text.lower() == "nan" else text


def ticket_type_bucket(row: pd.Series) -> str:
    ticket_type = clean(row.get("TicketType", "")).upper()
    ticket_type_text = clean(row.get("TicketTypeText", "")).lower()
    if ticket_type == "Z010" or ticket_type_text == "pdi":
        return "PDI ticket"
    if ticket_type == "Z007" or ticket_type_text == "repair ticket":
        return "Repair ticket"
    return "Other"


def workflow_stage(row: pd.Series) -> str:
    status = clean(row.get("TicketStatusText", ""))
    key = status.lower()
    bucket = ticket_type_bucket(row)
    if bucket == "PDI ticket":
        return {
            "open": "Awaiting PDI Start",
            "repair completed": "PDI Complete, Awaiting Time Claim",
            "claim time ticket": "PDI Completed",
            "cancel": "Cancelled",
        }.get(key, status or "Blank")
    if bucket == "Repair ticket":
        return {
            "open": "Awaiting Quote Approval",
            "quote approved": "Approved, Awaiting Repair",
            "repair in progress": "Repair In Progress",
            "create delivery note": "Repair In Progress",
            "repair completed": "Repair Complete, Awaiting Time Claim",
            "claim time ticket": "Time Claimed, Awaiting Invoice",
            "create invoice": "Completed / Invoiced",
            "cancel invoice": "Invoice Cancelled, SO Still Open",
            "cancel": "Cancelled",
        }.get(key, status or "Blank")
    return status or "Blank"


def is_abnormal_party(row: pd.Series) -> bool:
    party = clean(row.get("Role_1001_InvolvedPartyName", "")).upper()
    party_id = clean(row.get("Role_1001_InvolvedPartyID", "")).upper()
    if not party:
        return False
    return party != INTERNAL_PARTY_NAME and party_id.endswith(".ORG")


def load_dashboard() -> dict[str, Any]:
    text = DATA_JS.read_text(encoding="utf-8")
    return json.loads(text[text.find("{"):text.rfind("}") + 1])


def expected_counts(rows: pd.DataFrame, period: str, ticket_type: str, dealer: str) -> dict[str, int]:
    selected = rows[
        rows["DealerBucket"].eq(dealer)
        & rows["TicketTypeBucket"].eq(ticket_type)
    ].copy()
    if period.isdigit():
        selected = selected[selected["CreatedYear"].eq(period)]
    else:
        selected = selected[selected["CreatedLabel"].eq(period)]
    counts = selected["WorkflowStage"].value_counts().to_dict()
    stage_order = REPAIR_STAGES if ticket_type == "Repair ticket" else PDI_STAGES
    return {stage: int(counts.get(stage, 0)) for stage in stage_order}


def dashboard_counts(data: dict[str, Any], period: str, ticket_type: str, dealer: str) -> dict[str, int]:
    node = data["pages"]["overview"]["workflowDaily"].get(period, {}).get(ticket_type, {}).get(dealer, {})
    return {row["status"]: int(row.get("qty", 0)) for row in node.get("pipeline", [])}


def main() -> None:
    data = load_dashboard()
    tickets = pd.read_excel(SOURCE, sheet_name="Tickets", dtype=str).fillna("")
    created = pd.to_datetime(tickets["CreatedOn"], errors="coerce", format="mixed", dayfirst=True)
    tickets["CreatedDate"] = created
    tickets["CreatedYear"] = created.dt.year.astype("Int64").astype(str).where(created.notna(), "")
    tickets["CreatedLabel"] = created.dt.strftime("%b %Y").where(created.notna(), "")
    tickets["DealerBucket"] = tickets["DealerName"].map(lambda value: DEALER_LABELS.get(clean(value), "Other"))
    tickets["TicketTypeBucket"] = tickets.apply(ticket_type_bucket, axis=1)
    tickets["WorkflowStage"] = tickets.apply(workflow_stage, axis=1)
    normal = tickets[~tickets.apply(is_abnormal_party, axis=1)].copy()

    periods = list(data["pages"]["overview"]["workflowDaily"].keys())
    mismatches = []
    totals_checked = 0
    for period in periods:
        for ticket_type in TYPES:
            for dealer in DEALERS:
                expected = expected_counts(normal, period, ticket_type, dealer)
                actual = dashboard_counts(data, period, ticket_type, dealer)
                totals_checked += 1
                for stage, expected_qty in expected.items():
                    actual_qty = int(actual.get(stage, 0))
                    if actual_qty != expected_qty:
                        mismatches.append(
                            {
                                "period": period,
                                "dealer": dealer,
                                "ticketType": ticket_type,
                                "stage": stage,
                                "expected": expected_qty,
                                "actual": actual_qty,
                            }
                        )

    spotlight = defaultdict(dict)
    for period in ["2026", "Aug 2026"]:
        for ticket_type in TYPES:
            spotlight[period][ticket_type] = dashboard_counts(data, period, ticket_type, "Perth")

    report = {
        "source": str(SOURCE),
        "dashboardData": str(DATA_JS),
        "lastUpdated": data["meta"]["lastUpdated"],
        "checks": totals_checked,
        "mismatchCount": len(mismatches),
        "mismatches": mismatches[:200],
        "perthSpotlight": spotlight,
    }
    OUT.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
