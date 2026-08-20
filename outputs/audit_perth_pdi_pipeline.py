from __future__ import annotations

from pathlib import Path

import pandas as pd


SOURCE = Path(r"C:\Users\Leo.Li\Documents\ChatGPT\Service centre dashboard\c4c_ticket_table_z007_z010_checked_hana_final.xlsx")
DEALER_LABELS = {
    "Regent RV - Perth": "Perth",
    "Regent RV - Traralgon": "Traralgon",
    "Snowy River Launceston": "Launceston",
    "Snowy River Geelong": "Geelong",
    "Regent RV - Frankston": "Frankston",
}


def clean(value: object) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    return "" if text.lower() == "nan" else text


def is_pdi(row: pd.Series) -> bool:
    return clean(row.get("TicketType", "")).upper() == "Z010" or clean(row.get("TicketTypeText", "")).lower() == "pdi"


def is_pdi_open(row: pd.Series) -> bool:
    return clean(row.get("TicketStatusText", "")).lower() not in {"claim time ticket", "cancel"}


def workflow_stage(status: str) -> str:
    return {
        "open": "Awaiting PDI Start",
        "repair completed": "PDI Complete, Awaiting Time Claim",
        "claim time ticket": "PDI Completed",
        "cancel": "Cancelled",
    }.get(clean(status).lower(), clean(status) or "Blank")


def show(name: str, rows: pd.DataFrame) -> None:
    print(f"\n{name}: {len(rows)}")
    if rows.empty:
        return
    rows = rows.copy()
    rows["WorkflowStage"] = rows["TicketStatusText"].map(workflow_stage)
    print("Stage counts:")
    print(rows["WorkflowStage"].value_counts(dropna=False).to_string())
    print("Raw status counts:")
    print(rows["TicketStatusText"].replace("", "Blank").value_counts(dropna=False).to_string())
    cols = [
        "TicketID",
        "CreatedOn",
        "TicketStatusText",
        "WorkflowStage",
        "DealerName",
        "Role_40_InvolvedPartyName",
        "TotalLabourHours",
        "SerialID",
        "ChassisNumber",
    ]
    existing = [col for col in cols if col in rows.columns]
    print(rows.sort_values(["WorkflowStage", "CreatedDate", "TicketID"])[existing].to_string(index=False))


def main() -> None:
    tickets = pd.read_excel(SOURCE, sheet_name="Tickets", dtype=str).fillna("")
    created = pd.to_datetime(tickets["CreatedOn"], errors="coerce", format="mixed", dayfirst=True)
    tickets["CreatedDate"] = created
    tickets["CreatedYear"] = created.dt.year.astype("Int64").astype(str).where(created.notna(), "")
    tickets["CreatedMonth"] = created.dt.to_period("M").astype(str).where(created.notna(), "")
    tickets["DealerBucket"] = tickets["DealerName"].map(lambda value: DEALER_LABELS.get(clean(value), "Other"))
    tickets["IsPDI"] = tickets.apply(is_pdi, axis=1)
    tickets["IsOpenTicket"] = tickets.apply(lambda row: is_pdi_open(row) if row["IsPDI"] else True, axis=1)

    base = tickets[tickets["DealerBucket"].eq("Perth") & tickets["IsPDI"]].copy()
    year_end = pd.Period("2026", freq="Y").to_timestamp(how="end")
    web_year_open = base[
        base["CreatedMonth"].ne("")
        & base["IsOpenTicket"]
        & base["CreatedDate"].le(year_end)
    ].copy()
    created_2026 = base[base["CreatedYear"].eq("2026")].copy()
    created_2026_open = created_2026[created_2026["IsOpenTicket"]].copy()

    show("WEB CURRENT YEAR OPEN PDI PERTH (current pipeline source)", web_year_open)
    show("CREATED IN 2026 ONLY + STILL OPEN PDI PERTH", created_2026_open)
    show("ALL CREATED IN 2026 PDI PERTH", created_2026)


if __name__ == "__main__":
    main()
