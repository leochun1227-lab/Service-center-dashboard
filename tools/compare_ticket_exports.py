from __future__ import annotations

import json
from pathlib import Path

import pandas as pd


FILES = {
    "desktop_ref": Path(r"C:\Users\Leo.Li\Desktop\c4c_ticket_table_z007_z010_with_invoice_layout_checked.xlsx"),
    "web_source": Path(r"C:\Users\Leo.Li\Documents\ChatGPT\Service centre dashboard\c4c_ticket_table_z007_z010_checked_hana_final.xlsx"),
    "latest_export": Path(
        r"C:\Users\Leo.Li\Documents\GitHub\Service-center-dashboard\outputs\01a01c4f-25b7-7e91-8ffd-41a61b3fc906\c4c_ticket_table_z007_z010_with_invoice_layout_checked_latest_like_reference.xlsx"
    ),
}

KEY_COLUMNS = [
    "TicketType",
    "TicketTypeText",
    "DealerID",
    "DealerName",
    "ERPInvoiceNumber",
    "ERPInvoiceNumberPrice",
    "Billing date",
    "AmountIncludingTax",
    "WarrantyHandlingDealerID",
    "CreatedOn",
    "TicketStatus",
    "TicketStatusText",
    "ChangeOnDateTime",
    "Role_1001_InvolvedPartyID",
    "Role_1001_InvolvedPartyName",
    "Role_40_InvolvedPartyName",
    "Z1Z8TimeConsumed",
    "TotalLabourHours",
]


def norm_id(series: pd.Series) -> pd.Series:
    return series.astype(str).str.strip().str.replace(r"\.0$", "", regex=True)


def norm_text(series: pd.Series) -> pd.Series:
    return series.fillna("").astype(str).str.strip()


def load() -> dict[str, pd.DataFrame]:
    return {name: pd.read_excel(path, sheet_name="Tickets", dtype=str).fillna("") for name, path in FILES.items()}


def file_summary(name: str, path: Path, frame: pd.DataFrame) -> dict:
    return {
        "path": str(path),
        "size": path.stat().st_size,
        "mtimeSydney": pd.Timestamp(path.stat().st_mtime, unit="s", tz="UTC")
        .tz_convert("Australia/Sydney")
        .strftime("%Y-%m-%d %H:%M:%S %Z"),
        "rows": int(len(frame)),
        "cols": int(len(frame.columns)),
        "uniqueTicketIDs": int(norm_id(frame["TicketID"]).nunique()),
        "hasTotalLabourHours": "TotalLabourHours" in frame.columns,
        "first10Cols": list(frame.columns[:10]),
        "statusCounts": frame["TicketStatusText"].value_counts().to_dict(),
        "typeCounts": frame["TicketTypeText"].value_counts().to_dict(),
    }


def id_set(frame: pd.DataFrame) -> set[str]:
    return set(norm_id(frame["TicketID"]))


def diff_common(frames: dict[str, pd.DataFrame], left: str, right: str) -> dict[str, int]:
    left_frame = frames[left].copy()
    right_frame = frames[right].copy()
    left_frame["_id"] = norm_id(left_frame["TicketID"])
    right_frame["_id"] = norm_id(right_frame["TicketID"])
    left_index = left_frame.set_index("_id")
    right_index = right_frame.set_index("_id")
    ids = sorted(set(left_index.index) & set(right_index.index))
    result = {}
    for column in KEY_COLUMNS:
        if column in left_index.columns and column in right_index.columns:
            left_values = norm_text(left_index.loc[ids, column])
            right_values = norm_text(right_index.loc[ids, column])
            result[column] = int((left_values != right_values).sum())
    return result


def sample_missing(frames: dict[str, pd.DataFrame], base: str, compare: str, limit: int = 20) -> list[dict]:
    base_frame = frames[base].copy()
    compare_ids = id_set(frames[compare])
    base_frame["_id"] = norm_id(base_frame["TicketID"])
    missing = base_frame[~base_frame["_id"].isin(compare_ids)].copy()
    cols = ["TicketID", "TicketTypeText", "DealerName", "TicketStatusText", "CreatedOn", "AmountIncludingTax"]
    return missing[[col for col in cols if col in missing.columns]].head(limit).to_dict(orient="records")


def main() -> None:
    frames = load()
    output = {name: file_summary(name, path, frames[name]) for name, path in FILES.items()}
    sets = {name: id_set(frame) for name, frame in frames.items()}
    output["idCompare"] = {}
    names = list(FILES)
    for idx, left in enumerate(names):
        for right in names[idx + 1 :]:
            output["idCompare"][f"{left}_vs_{right}"] = {
                f"{left}_not_in_{right}": int(len(sets[left] - sets[right])),
                f"{right}_not_in_{left}": int(len(sets[right] - sets[left])),
                "common": int(len(sets[left] & sets[right])),
                "sameIds": sets[left] == sets[right],
            }
    output["commonValueDiffs"] = {
        "desktop_ref_vs_web_source": diff_common(frames, "desktop_ref", "web_source"),
        "web_source_vs_latest_export": diff_common(frames, "web_source", "latest_export"),
    }
    output["sampleCurrentSourceNotInDesktop"] = sample_missing(frames, "web_source", "desktop_ref")
    output["sampleExportNotInDesktop"] = sample_missing(frames, "latest_export", "desktop_ref")
    print(json.dumps(output, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
