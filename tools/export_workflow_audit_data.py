from __future__ import annotations

import importlib.util
import json
from pathlib import Path
from typing import Any

import pandas as pd


DASHBOARD_DIR = Path(r"C:\Users\Leo.Li\Documents\ChatGPT\Service centre dashboard")
GENERATOR_PATH = DASHBOARD_DIR / "generate_web_data_from_excel.py"
SOURCE_WORKBOOK = DASHBOARD_DIR / "c4c_ticket_table_z007_z010_checked_hana_final.xlsx"
DASHBOARD_DATA = DASHBOARD_DIR / "dashboard-data.js"
OUTPUT_JSON = Path(r"C:\Users\Leo.Li\Documents\GitHub\Service-center-dashboard\outputs\workflow_audit_data.json")


def load_generator():
    spec = importlib.util.spec_from_file_location("dashboard_generator", GENERATOR_PATH)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Unable to load generator: {GENERATOR_PATH}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def load_dashboard_payload() -> dict[str, Any]:
    text = DASHBOARD_DATA.read_text(encoding="utf-8").strip()
    prefix = "window.serviceCentreData = "
    if not text.startswith(prefix):
        raise ValueError("dashboard-data.js does not start with expected assignment")
    if text.endswith(";"):
        text = text[:-1]
    return json.loads(text[len(prefix) :])


def period_to_month_code(label: str) -> str:
    return pd.Period(pd.to_datetime(label, format="%b %Y"), freq="M").strftime("%Y-%m")


def safe_float(series: pd.Series) -> pd.Series:
    return pd.to_numeric(
        series.fillna("").astype(str).str.replace(",", "", regex=False).str.replace("$", "", regex=False).str.strip(),
        errors="coerce",
    ).fillna(0.0)


def prepare_tickets(gen) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    tickets = pd.read_excel(SOURCE_WORKBOOK, sheet_name="Tickets", dtype=str).fillna("")
    created = pd.to_datetime(tickets["CreatedOn"], errors="coerce", format="mixed", dayfirst=True)
    billing = pd.to_datetime(tickets["Billing date"], errors="coerce", format="mixed", dayfirst=True)
    changed = pd.to_datetime(tickets["ChangeOnDateTime"], errors="coerce", format="mixed", dayfirst=True)
    tickets["CreatedDate"] = created
    tickets["BillingDate"] = billing
    tickets["ChangeDate"] = changed
    tickets["CreatedMonth"] = created.dt.to_period("M").astype(str).where(created.notna(), "")
    tickets["CreatedYear"] = created.dt.year.astype("Int64").astype(str).where(created.notna(), "")
    tickets["BillingMonth"] = billing.dt.to_period("M").astype(str).where(billing.notna(), "")
    tickets["BillingYear"] = billing.dt.year.astype("Int64").astype(str).where(billing.notna(), "")
    tickets["DealerBucket"] = tickets["DealerName"].map(gen.dealer_bucket)
    tickets["TicketTypeBucket"] = tickets.apply(gen.ticket_type_bucket, axis=1)
    tickets["WorkflowStage"] = tickets.apply(gen.workflow_stage, axis=1)
    tickets["RawTicketStatus"] = tickets["TicketStatusText"].map(lambda value: gen.clean(value) or "Blank")
    tickets["NewTicketQuoteAmount"] = gen.parse_amount(tickets["AmountIncludingTax"])
    tickets["InvoiceAmount"] = gen.parse_amount(tickets["ERPInvoiceNumberPrice"])
    tickets["LabourHours"] = gen.parse_hours(tickets["TotalLabourHours"]) if "TotalLabourHours" in tickets.columns else 0.0
    tickets["AbnormalParty"] = tickets.apply(gen.is_abnormal_party, axis=1)
    tickets["InvoiceScope"] = tickets.apply(gen.invoice_scope, axis=1)
    tickets["AbnormalReason"] = tickets.apply(lambda row: "; ".join(gen.abnormal_reasons(row)), axis=1)
    tickets["IncludedInWorkflow"] = ~tickets["AbnormalParty"]
    tickets["ExcludedFromWorkflowReason"] = tickets["AbnormalParty"].map(lambda value: "Abnormal party excluded by current Workflow logic" if value else "")
    tickets["IsOpenTicket"] = tickets.apply(gen.is_open_ticket, axis=1)

    normal_tickets = tickets[~tickets["AbnormalParty"]].copy()
    normal_tickets["CompletedDate"] = normal_tickets.apply(gen.completed_date, axis=1)
    open_tickets = tickets[tickets["CreatedMonth"].ne("") & tickets.apply(gen.is_open_ticket, axis=1)].copy()
    missing_invoice_mask = (
        normal_tickets["TicketStatusText"].map(lambda value: gen.clean(value).lower()).eq(gen.CREATE_INVOICE_STATUS)
        & (
            normal_tickets["ERPInvoiceNumber"].map(gen.clean).eq("")
            | normal_tickets["ERPInvoiceNumberPrice"].map(gen.clean).eq("")
        )
    )
    normal_tickets["MissingInvoicePlaceholder"] = missing_invoice_mask
    valid_invoice_mask = (
        normal_tickets["BillingMonth"].ne("")
        & normal_tickets["ERPInvoiceNumber"].map(gen.clean).ne("")
        & normal_tickets["ERPInvoiceNumberPrice"].map(gen.clean).ne("")
        & ~normal_tickets["MissingInvoicePlaceholder"]
    )
    invoice_tickets = normal_tickets[valid_invoice_mask].copy()
    invoice_tickets["InvoiceMonth"] = invoice_tickets["BillingMonth"]
    invoice_tickets["InvoiceYear"] = invoice_tickets["BillingYear"]
    return tickets, normal_tickets, invoice_tickets, open_tickets


def flatten_pipeline(gen, payload, normal_tickets, invoice_tickets, open_tickets) -> list[dict[str, Any]]:
    rows = []
    workflow_daily = payload["pages"]["overview"]["workflowDaily"]
    for period, type_map in workflow_daily.items():
        for ticket_type in ["Repair ticket", "PDI ticket"]:
            for dealer in [name for name in gen.DEALER_ORDER if name != "Other"]:
                dashboard_workflow = type_map.get(ticket_type, {}).get(dealer)
                if period.isdigit():
                    recalculated = gen.build_year_workflow(normal_tickets, invoice_tickets, open_tickets, period, ticket_type, dealer)
                    source_rule = f"CreatedYear = {period}; AbnormalParty excluded"
                else:
                    month_code = period_to_month_code(period)
                    recalculated = gen.build_daily_workflow(normal_tickets, invoice_tickets, open_tickets, month_code, ticket_type, dealer)
                    source_rule = f"CreatedMonth = {month_code}; AbnormalParty excluded"
                dash_by_stage = {
                    item["status"]: item
                    for item in (dashboard_workflow or {}).get("pipeline", [])
                }
                recalc_by_stage = {item["status"]: item for item in recalculated["pipeline"]}
                stages = gen.workflow_stages_for_type(ticket_type)
                for stage in stages:
                    dash = dash_by_stage.get(stage, {})
                    recalc = recalc_by_stage.get(stage, {})
                    rows.append(
                        {
                            "Period": period,
                            "Dealer": dealer,
                            "TicketType": ticket_type,
                            "WorkflowStage": stage,
                            "DashboardQty": int(dash.get("qty", 0)),
                            "RecalcQtyFromSource": int(recalc.get("qty", 0)),
                            "Difference": int(dash.get("qty", 0)) - int(recalc.get("qty", 0)),
                            "QuoteAmount": round(float(recalc.get("quoteAmount", 0)), 2),
                            "RawStatuses": ", ".join(recalc.get("rawStatuses", [])),
                            "SourceRule": source_rule,
                            "DashboardTotalCreated": int((dashboard_workflow or {}).get("totals", {}).get("created", 0)),
                            "RecalcTotalCreated": int(recalculated.get("totals", {}).get("created", 0)),
                            "DashboardOpen": int((dashboard_workflow or {}).get("totals", {}).get("open", 0)),
                            "RecalcOpen": int(recalculated.get("totals", {}).get("open", 0)),
                        }
                    )
    return rows


def detail_rows(tickets: pd.DataFrame) -> list[dict[str, Any]]:
    preferred = [
        "TicketID",
        "TicketType",
        "TicketTypeText",
        "DealerName",
        "DealerBucket",
        "CreatedOn",
        "CreatedMonth",
        "CreatedYear",
        "TicketStatusText",
        "RawTicketStatus",
        "WorkflowStage",
        "IncludedInWorkflow",
        "ExcludedFromWorkflowReason",
        "AbnormalParty",
        "AbnormalReason",
        "IsOpenTicket",
        "AmountIncludingTax",
        "NewTicketQuoteAmount",
        "ERPInvoiceNumber",
        "ERPInvoiceNumberPrice",
        "InvoiceAmount",
        "Billing date",
        "BillingMonth",
        "ChangeOnDateTime",
        "TotalLabourHours",
        "LabourHours",
        "Role_1001_InvolvedPartyID",
        "Role_1001_InvolvedPartyName",
        "Role_40_InvolvedPartyName",
        "Subject",
        "Description",
    ]
    columns = [col for col in preferred if col in tickets.columns]
    output = tickets[columns].copy()
    for col in ["CreatedOn", "Billing date", "ChangeOnDateTime"]:
        if col in output.columns:
            output[col] = output[col].map(lambda value: "" if pd.isna(value) else str(value))
    return output.to_dict(orient="records")


def mapping_rows() -> list[dict[str, Any]]:
    return [
        {"TicketType": "Repair ticket", "C4C Status": "Open", "Dashboard Display": "Awaiting Quote Approval"},
        {"TicketType": "Repair ticket", "C4C Status": "Quote Approved", "Dashboard Display": "Approved, Awaiting Repair"},
        {"TicketType": "Repair ticket", "C4C Status": "Repair in Progress", "Dashboard Display": "Repair In Progress"},
        {"TicketType": "Repair ticket", "C4C Status": "Create delivery note", "Dashboard Display": "Repair In Progress"},
        {"TicketType": "Repair ticket", "C4C Status": "Repair completed", "Dashboard Display": "Repair Complete, Awaiting Time Claim"},
        {"TicketType": "Repair ticket", "C4C Status": "Claim Time Ticket", "Dashboard Display": "Time Claimed, Awaiting Invoice"},
        {"TicketType": "Repair ticket", "C4C Status": "Create invoice", "Dashboard Display": "Completed / Invoiced"},
        {"TicketType": "Repair ticket", "C4C Status": "Cancel invoice", "Dashboard Display": "Invoice Cancelled, SO Still Open"},
        {"TicketType": "Repair ticket", "C4C Status": "Cancel", "Dashboard Display": "Cancelled"},
        {"TicketType": "PDI ticket", "C4C Status": "Open", "Dashboard Display": "Awaiting PDI Start"},
        {"TicketType": "PDI ticket", "C4C Status": "Repair completed", "Dashboard Display": "PDI Complete, Awaiting Time Claim"},
        {"TicketType": "PDI ticket", "C4C Status": "Claim Time Ticket", "Dashboard Display": "PDI Completed"},
        {"TicketType": "PDI ticket", "C4C Status": "Cancel", "Dashboard Display": "Cancelled"},
    ]


def main() -> None:
    gen = load_generator()
    payload = load_dashboard_payload()
    tickets, normal_tickets, invoice_tickets, open_tickets = prepare_tickets(gen)
    summary_rows = flatten_pipeline(gen, payload, normal_tickets, invoice_tickets, open_tickets)
    data = {
        "meta": {
            "sourceWorkbook": str(SOURCE_WORKBOOK),
            "dashboardData": str(DASHBOARD_DATA),
            "dashboardLastUpdated": payload["meta"]["lastUpdated"],
            "sourceRows": int(len(tickets)),
            "normalRowsIncludedInWorkflow": int(len(normal_tickets)),
            "abnormalPartyRowsExcluded": int(tickets["AbnormalParty"].sum()),
            "summaryRows": int(len(summary_rows)),
        },
        "workflowSummary": summary_rows,
        "workflowDetail": detail_rows(tickets),
        "statusMapping": mapping_rows(),
    }
    OUTPUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT_JSON.write_text(json.dumps(data, ensure_ascii=False, default=str), encoding="utf-8")
    print(json.dumps(data["meta"], ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
