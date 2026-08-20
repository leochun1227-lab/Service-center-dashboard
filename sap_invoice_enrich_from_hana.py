from __future__ import annotations

import os
from pathlib import Path
from typing import Iterable

import pandas as pd


WORKBOOK_PATH = Path(
    os.getenv(
        "WORKBOOK_PATH",
        r"C:\Users\Leo.Li\Documents\ChatGPT\Service centre dashboard\c4c_ticket_table_z007_z010_with_invoice_layout_checked.xlsx",
    )
)
OUTPUT_PATH = Path(
    os.getenv(
        "OUTPUT_PATH",
        r"C:\Users\Leo.Li\Documents\ChatGPT\Service centre dashboard\c4c_ticket_table_z007_z010_checked_hana_final.xlsx",
    )
)
TICKETS_SHEET = "Tickets"
NOT_ASSIGNED_SHEET = "NotAssigned"
RESULT_SHEET = "DealerMappingResult"
MAPPING_SHEET = "DealerMappingUsed"
SCHEMA = os.getenv("SAP_SCHEMA", "SAPHANADB")
SAP_CLIENT = os.getenv("SAP_CLIENT", "800")
SAP_HANA_DSN = os.getenv(
    "SAP_HANA_DSN",
    "DRIVER={HDBODBC};SERVERNODE=10.11.2.25:30241;UID=BAOJIANFENG;PWD=Xja@2025ABC;",
)


def clean(value: object) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if text.lower() == "nan":
        return ""
    return text


def chunked(values: list[str], size: int) -> Iterable[list[str]]:
    for i in range(0, len(values), size):
        yield values[i : i + size]


def load_hana_dsn() -> str:
    dsn = clean(SAP_HANA_DSN)
    if not dsn:
        raise SystemExit("SAP_HANA_DSN is empty. Set it before running SAP HANA enrichment.")
    return dsn


def fetch_invoice_map(invoice_numbers: list[str]) -> pd.DataFrame:
    import pyodbc

    dsn = load_hana_dsn()
    conn = pyodbc.connect(dsn, timeout=60, autocommit=True)
    rows: list[tuple[str, str, str]] = []
    try:
        cur = conn.cursor()
        for batch in chunked(invoice_numbers, 200):
            escaped = [n.replace("'", "''") for n in batch]
            in_list = ",".join(f"'{n}'" for n in escaped)
            sql = f"""
                SELECT
                    k."VBELN" AS "ERPInvoiceNumber",
                    TO_NVARCHAR(k."FKDAT", 'YYYY-MM-DD') AS "Billing date",
                    CAST(COALESCE(k."NETWR", 0) AS DECIMAL(15, 2)) AS "ERPInvoiceNumberPrice"
                FROM "{SCHEMA}"."VBRK" k
                WHERE k."MANDT" = '{SAP_CLIENT.replace("'", "''")}'
                  AND k."VBELN" IN ({in_list})
            """
            cur.execute(sql)
            for row in cur.fetchall():
                rows.append((clean(row[0]), clean(row[1]), clean(row[2])))
    finally:
        conn.close()

    return pd.DataFrame(rows, columns=["ERPInvoiceNumber", "Billing date", "ERPInvoiceNumberPrice"])


def enrich_sheet(df: pd.DataFrame, invoice_map: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out["ERPInvoiceNumber"] = out.get("ERPInvoiceNumber", "").fillna("").astype(str).str.strip()
    out = out.drop(columns=[c for c in ["ERPInvoiceNumberPrice", "Billing date"] if c in out.columns])
    out = out.merge(invoice_map, how="left", on="ERPInvoiceNumber")
    out["ERPInvoiceNumberPrice"] = out["ERPInvoiceNumberPrice"].fillna("")
    out["Billing date"] = out["Billing date"].fillna("")

    preferred = [
        "TicketID",
        "TicketType",
        "TicketTypeText",
        "DealerID",
        "DealerName",
        "ERPInvoiceNumber",
        "ERPInvoiceNumberPrice",
        "Billing date",
        "AmountIncludingTax",
        "TotalLabourHours",
        "WarrantyHandlingDealerID",
        "CreatedOn",
        "TicketStatus",
        "TicketStatusText",
        "ERPFreeOrder",
        "Role_40_InvolvedPartyName",
        "Role_43_InvolvedPartyName",
        "TicketName",
        "SerialID",
        "ChassisNumber",
        "DealerResolutionStatus",
        "OriginalDealerName",
    ]
    cols = [c for c in preferred if c in out.columns] + [c for c in out.columns if c not in preferred]
    return out[cols]


def main() -> None:
    tickets = pd.read_excel(WORKBOOK_PATH, sheet_name=TICKETS_SHEET, dtype=str).fillna("")
    not_assigned = pd.read_excel(WORKBOOK_PATH, sheet_name=NOT_ASSIGNED_SHEET, dtype=str).fillna("")
    result = pd.read_excel(WORKBOOK_PATH, sheet_name=RESULT_SHEET, dtype=str).fillna("")
    mapping = pd.read_excel(WORKBOOK_PATH, sheet_name=MAPPING_SHEET, dtype=str).fillna("")

    invoices = sorted({clean(v) for v in tickets["ERPInvoiceNumber"].tolist() if clean(v)})
    invoice_map = fetch_invoice_map(invoices) if invoices else pd.DataFrame(
        columns=["ERPInvoiceNumber", "Billing date", "ERPInvoiceNumberPrice"]
    )

    tickets_out = enrich_sheet(tickets, invoice_map)
    not_assigned_out = enrich_sheet(not_assigned, invoice_map)

    with pd.ExcelWriter(OUTPUT_PATH, engine="openpyxl") as writer:
        tickets_out.to_excel(writer, index=False, sheet_name=TICKETS_SHEET)
        not_assigned_out.to_excel(writer, index=False, sheet_name=NOT_ASSIGNED_SHEET)
        result.to_excel(writer, index=False, sheet_name=RESULT_SHEET)
        mapping.to_excel(writer, index=False, sheet_name=MAPPING_SHEET)
        invoice_map.to_excel(writer, index=False, sheet_name="SAPInvoiceLookup")

    print(f"Workbook written: {OUTPUT_PATH}")
    print(f"Invoice rows matched: {len(invoice_map)}")


if __name__ == "__main__":
    main()
