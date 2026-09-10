"""Regression checks for the C4C quote decision; --fixture emits browser test data."""
import json
from pathlib import Path
import sys
import unittest

sys.dont_write_bytecode = True
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import pandas as pd
import generate_web_data_from_excel as dashboard


def fixture_rows():
    statuses = [
        ("OPEN", "1", "Open"),
        ("APPROVED", "2", "Quote Approved"),
        ("REJECTED", "ZW", "Quote Rejected"),
        ("REPAIR", "Y2", "Repair in Progress"),
    ]
    return pd.DataFrame([
        {"TicketID": ticket_id, "TicketType": "Z007", "TicketStatus": code,
         "TicketStatusText": text, "NewTicketQuoteAmount": 100,
         "CreatedDate": pd.Timestamp("2026-09-01"), "LastChangedDate": pd.Timestamp("2026-09-02")}
        for ticket_id, code, text in statuses
    ])


class QuoteRejectedTests(unittest.TestCase):
    def test_code_and_label_variants(self):
        for code, text in [("ZW", ""), ("", "ZW"), ("", "ZW - Quote Rejected"),
                           ("", " quote rejected "), ("zw", "Quote Rejected")]:
            with self.subTest(code=code, text=text):
                row = pd.Series({"TicketType": "Z007", "TicketStatus": code, "TicketStatusText": text})
                self.assertEqual(dashboard.workflow_stage(row), "Quote Rejected")

    def test_rejection_is_separate_from_approved_and_repair(self):
        rows = fixture_rows()
        pipeline = dashboard.build_status_pipeline(rows, "Repair ticket")
        stages = {stage["status"]: stage for stage in pipeline}
        self.assertEqual(sum(stage["qty"] for stage in pipeline), 3)
        self.assertNotIn("Quote Rejected", stages)
        self.assertEqual(stages["Approved, Awaiting Repair"]["ticketIds"], ["APPROVED"])
        self.assertEqual(stages["Repair In Progress"]["ticketIds"], ["REPAIR"])
        self.assertEqual(stages["Completed / Invoiced"]["qty"], 0)
        self.assertTrue(pd.isna(dashboard.completed_date(rows.iloc[2])))
        mix = dashboard.build_open_status_mix(rows, "Repair ticket")
        self.assertEqual(sum(stage["qty"] for stage in mix["segments"]), 4)
        rejected = next(stage for stage in mix["segments"] if stage["name"] == "Quote Rejected")
        self.assertEqual(rejected["qty"], 1)
        self.assertEqual(rejected["share"], 25)
        self.assertEqual(rejected["color"], "#ef4444")
        self.assertEqual(rejected["quoteAmount"], 100)
        self.assertEqual(sum(bucket["qty"] for bucket in rejected["aging"]), 1)

    def test_zero_state_and_pdi_are_separate(self):
        empty = fixture_rows().iloc[:0]
        repair = dashboard.build_status_pipeline(empty, "Repair ticket")
        self.assertEqual(len(repair), 8)
        self.assertNotIn("Quote Rejected", [stage["status"] for stage in repair])
        pdi = dashboard.build_status_pipeline(empty, "PDI ticket")
        self.assertEqual(len(pdi), 4)
        self.assertNotIn("Quote Rejected", [stage["status"] for stage in pdi])

    def test_unknown_status_does_not_mutate_stage_definitions(self):
        rows = fixture_rows()
        rows.loc[0, "TicketStatusText"] = "Unknown test status"
        dashboard.build_status_pipeline(rows, "Repair ticket")
        self.assertNotIn("Unknown test status", dashboard.REPAIR_WORKFLOW_STAGES)


if __name__ == "__main__":
    if "--fixture" in sys.argv:
        rows = fixture_rows()
        print(json.dumps({"pipeline": dashboard.build_status_pipeline(rows, "Repair ticket"),
                          "mix": dashboard.build_open_status_mix(rows, "Repair ticket"),
                          "details": [{"ticketId": row["TicketID"], "status": dashboard.workflow_stage(row),
                                       "rawStatus": row["TicketStatusText"]} for _, row in rows.iterrows()]}))
    else:
        unittest.main()
