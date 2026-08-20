from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pandas as pd


SOURCE = Path(
    r"C:\Users\Leo.Li\Documents\ChatGPT\Service centre dashboard"
    r"\c4c_ticket_table_z007_z010_checked_hana_final.xlsx"
)
OUT = Path(__file__).with_name("labour_audit_data.json")

DEALER_LABELS = {
    "Regent RV - Perth": "Perth",
    "Regent RV - Traralgon": "Traralgon",
    "Snowy River Launceston": "Launceston",
    "Snowy River Geelong": "Geelong",
    "Regent RV - Frankston": "Frankston",
}
DEALER_ORDER = ["Perth", "Traralgon", "Launceston", "Geelong", "Frankston", "Other"]
AUDIT_MONTH = "2026-08"
AUDIT_MONTH_LABEL = "Aug 2026"


def clean(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if text.lower() == "nan":
        return ""
    return text


def parse_hours(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series.fillna("").astype(str).str.strip(), errors="coerce").fillna(0.0)


def dealer_bucket(dealer_name: str) -> str:
    return DEALER_LABELS.get(clean(dealer_name), "Other")


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
    status_key = status.lower()
    ticket_type = ticket_type_bucket(row)
    if ticket_type == "PDI ticket":
        return {
            "open": "Awaiting PDI Start",
            "repair completed": "PDI Complete, Awaiting Time Claim",
            "claim time ticket": "PDI Completed",
            "cancel": "Cancelled",
        }.get(status_key, status or "Blank")
    if ticket_type == "Repair ticket":
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
        }.get(status_key, status or "Blank")
    return status or "Blank"


def safe_value(value: Any) -> Any:
    if pd.isna(value):
        return None
    if isinstance(value, pd.Timestamp):
        return value.strftime("%Y-%m-%d")
    return value


def make_worker_summary(rows: pd.DataFrame, include_dealer: bool) -> pd.DataFrame:
    group_cols = ["DealerBucket", "WorkerName"] if include_dealer else ["WorkerName"]
    records = []
    for keys, group in rows.groupby(group_cols, dropna=False):
        if include_dealer:
            dealer, worker = keys
        else:
            dealer, worker = "All Dealers", keys
        repair = group[group["TicketTypeBucket"].eq("Repair ticket")]
        pdi = group[group["TicketTypeBucket"].eq("PDI ticket")]
        hours = float(group["LabourHours"].sum())
        ticket_count = int(len(group))
        nonzero = int((group["LabourHours"] != 0).sum())
        records.append(
            {
                "DealerBucket": clean(dealer),
                "WorkerName": clean(worker) or "Unassigned",
                "TicketCount": ticket_count,
                "NonZeroHourTickets": nonzero,
                "ZeroHourTickets": ticket_count - nonzero,
                "TotalLabourHours": round(hours, 4),
                "AvgHoursAllTickets": round(hours / ticket_count, 4) if ticket_count else 0,
                "AvgHoursNonZeroTickets": round(hours / nonzero, 4) if nonzero else 0,
                "RepairHours": round(float(repair["LabourHours"].sum()), 4),
                "PDIHours": round(float(pdi["LabourHours"].sum()), 4),
                "RepairTickets": int(len(repair)),
                "PDITickets": int(len(pdi)),
            }
        )
    summary = pd.DataFrame(records)
    if summary.empty:
        return summary
    return summary.sort_values(
        ["TotalLabourHours", "DealerBucket", "WorkerName"],
        ascending=[False, True, True],
    ).reset_index(drop=True)


def df_to_rows(df: pd.DataFrame) -> list[list[Any]]:
    if df.empty:
        return [list(df.columns)]
    values = [[safe_value(v) for v in row] for row in df.itertuples(index=False, name=None)]
    return [list(df.columns), *values]


def main() -> None:
    tickets = pd.read_excel(SOURCE, sheet_name="Tickets", dtype=str).fillna("")
    created = pd.to_datetime(tickets["CreatedOn"], errors="coerce", format="mixed", dayfirst=True)
    tickets["CreatedDate"] = created
    tickets["CreatedMonth"] = created.dt.to_period("M").astype(str).where(created.notna(), "")
    tickets["DealerBucket"] = tickets["DealerName"].map(dealer_bucket)
    tickets["TicketTypeBucket"] = tickets.apply(ticket_type_bucket, axis=1)
    tickets["WorkflowStage"] = tickets.apply(workflow_stage, axis=1)
    tickets["LabourHours"] = parse_hours(tickets["TotalLabourHours"]) if "TotalLabourHours" in tickets.columns else 0.0
    tickets["WorkerName"] = tickets.get("Role_40_InvolvedPartyName", pd.Series(dtype="object")).map(
        lambda value: clean(value) or "Unassigned"
    )

    audit = tickets[tickets["CreatedMonth"].eq(AUDIT_MONTH)].copy()
    perth = audit[audit["DealerBucket"].eq("Perth")].copy()
    all_detail_cols = [
        "TicketID",
        "TicketType",
        "TicketTypeText",
        "DealerName",
        "DealerBucket",
        "TicketTypeBucket",
        "TicketStatusText",
        "WorkflowStage",
        "CreatedOn",
        "CreatedMonth",
        "Role_40_InvolvedPartyName",
        "WorkerName",
        "TotalLabourHours",
        "LabourHours",
        "AmountIncludingTax",
        "ERPInvoiceNumber",
        "ERPInvoiceNumberPrice",
        "Billing date",
    ]
    existing_cols = [col for col in all_detail_cols if col in audit.columns]
    detail = audit[existing_cols].sort_values(["DealerBucket", "WorkerName", "TicketTypeBucket", "TicketID"])

    type_dealer = (
        audit.groupby(["DealerBucket", "TicketTypeBucket"], dropna=False)
        .agg(TicketCount=("TicketID", "size"), TotalLabourHours=("LabourHours", "sum"))
        .reset_index()
    )
    type_dealer["TotalLabourHours"] = type_dealer["TotalLabourHours"].round(4)
    type_dealer = type_dealer.sort_values(["DealerBucket", "TicketTypeBucket"])

    status_summary = (
        audit.groupby(["DealerBucket", "TicketTypeBucket", "WorkflowStage"], dropna=False)
        .agg(TicketCount=("TicketID", "size"), TotalLabourHours=("LabourHours", "sum"))
        .reset_index()
    )
    status_summary["TotalLabourHours"] = status_summary["TotalLabourHours"].round(4)
    status_summary = status_summary.sort_values(["DealerBucket", "TicketTypeBucket", "WorkflowStage"])

    dashboard_check = []
    for dealer in DEALER_ORDER:
        dealer_rows = audit[audit["DealerBucket"].eq(dealer)]
        repair = dealer_rows[dealer_rows["TicketTypeBucket"].eq("Repair ticket")]
        pdi = dealer_rows[dealer_rows["TicketTypeBucket"].eq("PDI ticket")]
        dashboard_check.append(
            {
                "WebMonthLabel": AUDIT_MONTH_LABEL,
                "CreatedMonth": AUDIT_MONTH,
                "DealerBucket": dealer,
                "TicketCount": int(len(dealer_rows)),
                "Workers": int(dealer_rows.loc[dealer_rows["WorkerName"].ne("Unassigned"), "WorkerName"].nunique()),
                "TotalLabourHours": round(float(dealer_rows["LabourHours"].sum()), 4),
                "RepairHours": round(float(repair["LabourHours"].sum()), 4),
                "PDIHours": round(float(pdi["LabourHours"].sum()), 4),
            }
        )
    dashboard_check_df = pd.DataFrame(dashboard_check)

    payload = {
        "source": str(SOURCE),
        "auditMonth": AUDIT_MONTH,
        "sheets": {
            "Dashboard_Check": df_to_rows(dashboard_check_df),
            "Perth_Workers_Aug_2026": df_to_rows(make_worker_summary(perth, include_dealer=True)),
            "All_Workers_Aug_2026": df_to_rows(make_worker_summary(audit, include_dealer=True)),
            "Dealer_Type_Aug_2026": df_to_rows(type_dealer),
            "Status_Aug_2026": df_to_rows(status_summary),
            "Ticket_Detail_Aug_2026": df_to_rows(detail),
        },
    }
    OUT.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    print(json.dumps({"out": str(OUT), "rows": len(audit), "perthRows": len(perth)}, ensure_ascii=False))


if __name__ == "__main__":
    main()
