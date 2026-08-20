from __future__ import annotations

import json
from pathlib import Path

import pandas as pd


REFERENCE_WORKBOOK = Path(r"C:\Users\Leo.Li\Desktop\c4c_ticket_table_z007_z010_with_invoice_layout_checked.xlsx")
CURRENT_WORKBOOK = Path(r"C:\Users\Leo.Li\Documents\ChatGPT\Service centre dashboard\c4c_ticket_table_z007_z010_checked_hana_final.xlsx")
OUTPUT_JSON = Path(r"C:\Users\Leo.Li\Documents\GitHub\Service-center-dashboard\outputs\latest_ticket_export_like_reference.json")


def clean_cell(value):
    if pd.isna(value):
        return ""
    if isinstance(value, pd.Timestamp):
        return value.strftime("%Y-%m-%d %H:%M:%S") if value.time() != pd.Timestamp(value.date()).time() else value.strftime("%Y-%m-%d")
    return value


def reference_columns(sheet_name: str) -> list[str]:
    frame = pd.read_excel(REFERENCE_WORKBOOK, sheet_name=sheet_name, dtype=str, nrows=0)
    columns = list(frame.columns)
    if sheet_name in {"Tickets", "NotAssigned"} and "TotalLabourHours" not in columns:
        insert_after = "AmountIncludingTax"
        idx = columns.index(insert_after) + 1 if insert_after in columns else len(columns)
        columns = [*columns[:idx], "TotalLabourHours", *columns[idx:]]
    return columns


def load_sheet(sheet_name: str) -> dict:
    current = pd.read_excel(CURRENT_WORKBOOK, sheet_name=sheet_name, dtype=str).fillna("")
    columns = reference_columns(sheet_name)
    for column in columns:
        if column not in current.columns:
            current[column] = ""
    current = current[columns]
    rows = current.map(clean_cell).values.tolist()
    return {
        "name": sheet_name,
        "columns": columns,
        "rows": rows,
        "rowCount": int(len(rows)),
        "columnCount": int(len(columns)),
    }


def main() -> None:
    sheets = []
    for sheet_name in ["Tickets", "DealerMappingResult", "NotAssigned", "DealerMappingUsed"]:
        sheets.append(load_sheet(sheet_name))
    payload = {
        "sourceWorkbook": str(CURRENT_WORKBOOK),
        "referenceWorkbook": str(REFERENCE_WORKBOOK),
        "sheets": sheets,
    }
    OUTPUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT_JSON.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    print(
        json.dumps(
            {
                "sourceWorkbook": payload["sourceWorkbook"],
                "referenceWorkbook": payload["referenceWorkbook"],
                "sheets": [
                    {"name": sheet["name"], "rows": sheet["rowCount"], "cols": sheet["columnCount"]}
                    for sheet in sheets
                ],
            },
            ensure_ascii=False,
            indent=2,
        )
    )


if __name__ == "__main__":
    main()
