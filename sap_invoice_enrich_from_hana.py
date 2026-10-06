from __future__ import annotations

import os
from pathlib import Path
from typing import Iterable

import pandas as pd
from dashboard_credentials import apply_saved_settings

apply_saved_settings()


BASE_DIR = Path(__file__).resolve().parent
WORKBOOK_PATH = Path(
    os.getenv(
        "WORKBOOK_PATH",
        str(BASE_DIR / "c4c_ticket_table_z007_z010_with_invoice_layout_checked.xlsx"),
    )
)
OUTPUT_PATH = Path(
    os.getenv(
        "OUTPUT_PATH",
        str(BASE_DIR / "c4c_ticket_table_z007_z010_checked_hana_final.xlsx"),
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


def normalize_invoice_number(value: object) -> str:
    """Return the SAP billing document key in its 10-digit form."""
    text = clean(value)
    if text.isdigit():
        return text.zfill(10)
    return text


def chunked(values: list[str], size: int) -> Iterable[list[str]]:
    for i in range(0, len(values), size):
        yield values[i : i + size]


def load_hana_dsn() -> str:
    dsn = clean(SAP_HANA_DSN)
    if not dsn:
        raise SystemExit("SAP_HANA_DSN is empty. Set it before running SAP HANA enrichment.")
    return dsn


INVOICE_COLUMNS = [
    "ERPInvoiceNumber",
    "Billing date",
    "ERPInvoiceNumberPrice",
    "ERPInvoiceNumberPriceRaw",
    "BillingType",
]


def _invoice_frame(rows: list[tuple[object, object, object, object]]) -> pd.DataFrame:
    data = []
    for invoice, billing_date, netwr, billing_type in rows:
        raw = pd.to_numeric(pd.Series([clean(netwr)]), errors="coerce").fillna(0).iloc[0]
        doc_type = clean(billing_type).upper()
        # VBRK.NETWR is positive for cancellation documents.  The SAP
        # service-invoice total uses cancellation documents as credits.
        signed = -float(raw) if doc_type in {"S1", "AB"} else float(raw)
        data.append(
            (
                normalize_invoice_number(invoice),
                clean(billing_date),
                f"{signed:.2f}",
                f"{float(raw):.2f}",
                doc_type,
            )
        )
    return pd.DataFrame(data, columns=INVOICE_COLUMNS).drop_duplicates(
        subset=["ERPInvoiceNumber"], keep="first"
    )


def fetch_invoice_map(invoice_numbers: list[str]) -> pd.DataFrame:
    import pyodbc

    dsn = load_hana_dsn()
    conn = pyodbc.connect(dsn, timeout=60, autocommit=True)
    rows: list[tuple[object, object, object, object]] = []
    try:
        cur = conn.cursor()
        for batch in chunked(invoice_numbers, 200):
            escaped = [n.replace("'", "''") for n in batch]
            in_list = ",".join(f"'{n}'" for n in escaped)
            sql = f"""
                SELECT
                    k."VBELN" AS "ERPInvoiceNumber",
                    TO_NVARCHAR(k."FKDAT", 'YYYY-MM-DD') AS "Billing date",
                    CAST(COALESCE(k."NETWR", 0) AS DECIMAL(15, 2)) AS "ERPInvoiceNumberPriceRaw",
                    k."FKART" AS "BillingType"
                FROM "{SCHEMA}"."VBRK" k
                WHERE k."MANDT" = '{SAP_CLIENT.replace("'", "''")}'
                  AND k."VBELN" IN ({in_list})
            """
            cur.execute(sql)
            for row in cur.fetchall():
                rows.append((row[0], row[1], row[2], row[3]))
    finally:
        conn.close()

    return _invoice_frame(rows)


def fetch_service_invoice_map(invoice_numbers: list[str]) -> pd.DataFrame:
    """Load the complete SAP service-invoice universe plus requested C4C IDs.

    The material filter is the source-of-truth for SAP-only invoices.  The
    explicit ID query keeps C4C invoices visible even when their material is
    not present in the exported VBRP rows.
    """
    import pyodbc

    dsn = load_hana_dsn()
    conn = pyodbc.connect(dsn, timeout=60, autocommit=True)
    rows: list[tuple[object, object, object, object]] = []
    try:
        cur = conn.cursor()
        service_sql = f'''
            SELECT DISTINCT
                k."VBELN" AS "ERPInvoiceNumber",
                TO_NVARCHAR(k."FKDAT", 'YYYY-MM-DD') AS "Billing date",
                CAST(COALESCE(k."NETWR", 0) AS DECIMAL(15, 2)) AS "ERPInvoiceNumberPriceRaw",
                k."FKART" AS "BillingType"
            FROM "{SCHEMA}"."VBRK" k
            INNER JOIN "{SCHEMA}"."VBRP" p ON p."VBELN" = k."VBELN"
            WHERE k."MANDT" = '{SAP_CLIENT.replace("'", "''")}'
              AND p."MANDT" = '{SAP_CLIENT.replace("'", "''")}'
              AND p."MATNR" IN ('SOWTY999', 'SOWTY777')
        '''
        cur.execute(service_sql)
        rows.extend((row[0], row[1], row[2], row[3]) for row in cur.fetchall())

        normalized = sorted({normalize_invoice_number(value) for value in invoice_numbers if clean(value)})
        for batch in chunked(normalized, 200):
            escaped = [n.replace("'", "''") for n in batch]
            in_list = ",".join(f"'{n}'" for n in escaped)
            sql = f'''
                SELECT
                    k."VBELN" AS "ERPInvoiceNumber",
                    TO_NVARCHAR(k."FKDAT", 'YYYY-MM-DD') AS "Billing date",
                    CAST(COALESCE(k."NETWR", 0) AS DECIMAL(15, 2)) AS "ERPInvoiceNumberPriceRaw",
                    k."FKART" AS "BillingType"
                FROM "{SCHEMA}"."VBRK" k
                WHERE k."MANDT" = '{SAP_CLIENT.replace("'", "''")}'
                  AND k."VBELN" IN ({in_list})
            '''
            cur.execute(sql)
            rows.extend((row[0], row[1], row[2], row[3]) for row in cur.fetchall())
    finally:
        conn.close()

    return _invoice_frame(rows)


def enrich_sheet(df: pd.DataFrame, invoice_map: pd.DataFrame) -> pd.DataFrame:
    out = df.drop(columns=[c for c in ["ChangeOnDateTime"] if c in df.columns]).copy()
    out["ERPInvoiceNumber"] = out.get("ERPInvoiceNumber", "").map(normalize_invoice_number)
    out = out.drop(columns=[c for c in INVOICE_COLUMNS[1:] if c in out.columns])
    lookup = invoice_map.copy()
    lookup["ERPInvoiceNumber"] = lookup["ERPInvoiceNumber"].map(normalize_invoice_number)
    out = out.merge(lookup, how="left", on="ERPInvoiceNumber")
    for column in INVOICE_COLUMNS[1:]:
        out[column] = out[column].fillna("")

    preferred = [
        "TicketID",
        "TicketType",
        "TicketTypeText",
        "DealerID",
        "DealerName",
        "ERPInvoiceNumber",
        "ERPInvoiceNumberPrice",
        "ERPInvoiceNumberPriceRaw",
        "Billing date",
        "BillingType",
        "AmountIncludingTax",
        "TotalLabourHours",
        "WarrantyHandlingDealerID",
        "CreatedOn",
        "lastchangedtime",
        "TicketStatus",
        "TicketStatusText",
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

    invoice_values = [*tickets["ERPInvoiceNumber"].tolist(), *not_assigned["ERPInvoiceNumber"].tolist()]
    invoices = sorted({normalize_invoice_number(v) for v in invoice_values if clean(v)})
    invoice_map = fetch_service_invoice_map(invoices) if invoices else pd.DataFrame(columns=INVOICE_COLUMNS)

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
