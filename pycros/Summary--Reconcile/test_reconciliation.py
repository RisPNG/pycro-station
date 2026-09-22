import importlib.util
import sys
import unittest
from pathlib import Path
from unittest.mock import patch


spec = importlib.util.spec_from_file_location("summary_reconcile", Path(__file__).with_name("main.py"))
recon = importlib.util.module_from_spec(spec)
sys.modules[spec.name] = recon
spec.loader.exec_module(recon)


class MovementReconciliationTests(unittest.TestCase):
    def test_full_early_movements_preserve_original_amount_in_one_row(self):
        for job, qty, bds_amount, moved_amount in [
            ("BJ131190LS", 3439, 28646.87, 28543.70),
            ("BJ131191LS", 3138, 26139.54, 26045.40),
            ("AH001000LS", 10, 100, 110),
        ]:
            with self.subTest(job=job):
                rows = [(job, qty, bds_amount, 0, 0, "")]
                remark = "Early ship fr Oct'26 to Sep'26"
                self.assertTrue(recon._allocate_early_shipment(rows, job, qty, moved_amount, remark))
                self.assertEqual(rows, [(job, qty, bds_amount, 0, 0, remark)])

    def test_partial_delay_carries_net_amount_not_forecast_price(self):
        months = ["2026-08", "2026-09", "2026-10"]
        job = "BH082022MJ"
        rows = {month: [] for month in months}
        rows[months[0]] = [(job, 568, 19993.60, 561, 19360.11, "Missing")]
        shipment = {month: [] for month in months}
        shipment[months[2]] = [recon.SourceLine(job, job, 7, 250.81, "", "", "Forecast", 1)]
        ann = recon._new_month_maps(months)
        recon._reconcile_missing_movements(rows, months, {"2026-07": {}}, ann, shipment, "2026-07", [], 0.5)
        self.assertEqual(rows[months[2]][0][3], 7)
        self.assertAlmostEqual(rows[months[2]][0][4], 633.49)
        self.assertAlmostEqual(ann[months[2]][job][1], 633.49)
        self.assertEqual(rows[months[0]][0][:5], (job, 568, 19993.60, 561, 19360.11))

    def test_early_allocations_use_only_unmarked_remainder(self):
        job = "BF285013MJ"
        rows = [(job, 1102, 18028.72, 0, 0, "")]
        first = "Early ship fr Oct'26 to Aug'26"
        second = "Early ship fr Oct'26 to Sep'26"
        self.assertTrue(recon._allocate_early_shipment(rows, job, 252, 3678.63, first))
        self.assertFalse(recon._allocate_early_shipment(rows, job, 851, 14350.09, second))
        self.assertTrue(recon._allocate_early_shipment(rows, job, 850, 14350.09, second))
        self.assertEqual(len(rows), 2)
        self.assertEqual([row[-1] for row in rows], [first, second])
        self.assertEqual(sum(row[1] for row in rows), 1102)
        self.assertAlmostEqual(sum(row[2] for row in rows), 18028.72)

    def test_forecast_range_extends_to_fiscal_april(self):
        wb = recon.Workbook()
        ws = wb.active
        ws.title = "SHIPMENTS"
        ws.append(["JOB", "QTY", "AMOUNT", "PLAN EX-FTY"])
        ws.append(["AH001000LS", 10, 100, recon.date(2026, 8, 1)])
        with patch.object(recon, "load_workbook", return_value=wb):
            shipment, statuses, counts, months, warnings = recon._read_shipment_forecast("forecast.xlsx", lambda message: None)
        self.assertEqual(months, recon._month_range("2026-08", "2027-04"))
        self.assertEqual(counts["2027-04"], 0)
        self.assertEqual(len(shipment["2026-08"]), 1)

    def test_unmatched_quantity_stays_missing(self):
        months = ["2026-08", "2026-09"]
        job = "AU864564MJ"
        rows = {months[0]: [(job, 0, 0, 2, 20.12, "Missing")], months[1]: []}
        weekly = [recon.SourceLine(job, job, 330, 3326.4, "", "", "July", 1)]
        recon._reconcile_missing_movements(
            rows, months, {"2026-07": {job: [330, 3300]}},
            recon._new_month_maps(months), {month: [] for month in months},
            "2026-07", weekly, 0.5,
        )
        self.assertEqual(rows[months[0]][0][-1], "Missing")

    def test_existing_early_match_marks_later_month(self):
        months = ["2026-08", "2026-09", "2026-10"]
        job = "AH001000LS"
        bds = recon._new_month_maps(["2026-07", *months])
        ann = recon._new_month_maps(months)
        bds[months[2]][job] = [10, 100]
        ann[months[0]][job] = [10, 110]
        rows = {month: recon._build_reconciliation_rows(month, months, bds, ann, {}, 0.5, 0.01) for month in months}
        recon._reconcile_missing_movements(rows, months, bds, ann, {month: [] for month in months}, "2026-07", [], 0.5)
        self.assertEqual(rows[months[0]][0][-1], "Early ship fr Oct'26 to Aug'26")
        self.assertEqual(rows[months[2]][0][-1], rows[months[0]][0][-1])
        self.assertEqual(rows[months[2]][0][1:3], (10, 100))
        self.assertEqual(len(rows[months[2]]), 1)

    def test_summary_uses_full_totals_and_retains_manual_fx(self):
        months = ["2026-08", "2026-09", "2026-10"]
        bds = recon._new_month_maps(months)
        ann = recon._new_month_maps(months)
        bds[months[2]]["AH001000LS"] = [10, 100]
        rows = {month: recon._build_reconciliation_rows(month, months, bds, ann, {}, 0.5, 0.01) for month in months}
        wb = recon.Workbook()
        recon._write_summary_sheet(wb, recon.datetime(2026, 9, 21), months, bds, ann, rows)
        ws = wb["Summary FYE 2027"]
        self.assertEqual(ws["C8"].value, "='Recon Oct''26'!$C$1")
        self.assertTrue(ws["D8"].value.startswith("='Recon Oct''26'!$E$1-"))
        fx_cell = ws["D8"].value.split("-")[-1].replace("$", "")
        self.assertEqual(ws[fx_cell].value, 0)
        self.assertEqual(ws["E8"].value, "=C8-D8")

    def test_delayed_balance_excludes_quantity_already_shipped(self):
        months = ["2026-08", "2026-09"]
        job = "AU699028MJ"
        rows = {months[0]: [(job, 30, 528.3, 3190, 56526.5, "Missing")], months[1]: []}
        bds = {"2026-07": {job: [10169, 179076.09]}}
        weekly = [recon.SourceLine(job, job, 7009, 124199.48, "", "", "July", 1)]
        recon._reconcile_missing_movements(rows, months, bds, recon._new_month_maps(months), {month: [] for month in months}, "2026-07", weekly, 0.5)
        self.assertEqual(rows[months[0]][0][-1], "Delay ship fr Jul'26 to Aug'26")

    def test_historical_bds_quantity_matches_despite_price_difference(self):
        months = ["2026-08", "2026-09", "2026-10"]
        rows = {month: [] for month in months}
        rows[months[0]] = [("AK008021LJ", 0, 0, 9953, 99032.35, "Missing")]
        bds = {"2026-07": {"AK008021LJ": [9953, 98435.17]}}
        recon._reconcile_missing_movements(
            rows, months, bds, recon._new_month_maps(months),
            {month: [] for month in months}, "2026-07", [], 0.5,
        )
        self.assertEqual(rows[months[0]][0][-1], "Delay ship fr Jul'26 to Aug'26")
        self.assertEqual(rows[months[0]][0][4], 99032.35)

    def test_future_delay_uses_source_balance_and_consumes_rows(self):
        months = ["2026-08", "2026-09", "2026-10"]
        job = "BJ119062MJ"
        rows = {month: [] for month in months}
        rows[months[0]] = [(job, 402, 7364.64, 0, 0, "Missing")]
        rows[months[1]] = [(job, 402, 7364.64, 0, 0, "Missing")]
        shipment = {month: [] for month in months}
        shipment[months[2]] = [
            recon.SourceLine(job, job, 400, 6844, "", "", "Forecast", 1),
            recon.SourceLine(job + "-TOP", job, 2, 34.22, "", "", "Forecast", 2),
        ]
        ann = recon._new_month_maps(months)
        recon._reconcile_missing_movements(rows, months, {"2026-07": {}}, ann, shipment, "2026-07", [], 0.5)
        self.assertEqual(rows[months[2]], [(job, 0, 0, 402, 7364.64, "Delay ship fr Aug'26 to Oct'26")])
        self.assertEqual(rows[months[1]][0][-1], "Missing")
        self.assertEqual(ann[months[2]][job], [402, 7364.64])

    def test_early_matches_individual_line_and_preserves_future_bds(self):
        months = ["2026-08", "2026-09", "2026-10"]
        job = "BF285013MJ"
        rows = {month: [] for month in months}
        rows[months[1]] = [(job, 14551, 238054.36, 14803, 241732.99, "Missing")]
        rows[months[2]] = [(job, 1102, 18028.72, 0, 0, "")]
        shipment = {month: [] for month in months}
        shipment[months[2]] = [
            recon.SourceLine(job + "-11", job, 300, 4899, "", "", "Forecast", 1),
            recon.SourceLine(job + "-12", job, 252, 4115.16, "", "", "Forecast", 2),
        ]
        recon._reconcile_missing_movements(rows, months, {"2026-07": {}}, recon._new_month_maps(months), shipment, "2026-07", [], 0.5)
        self.assertEqual(rows[months[1]][0][-1], "Early ship fr Oct'26 to Sep'26")
        self.assertEqual(rows[months[2]][0][0:2], (job, 252))
        self.assertAlmostEqual(rows[months[2]][0][2], 3678.63)
        self.assertEqual(rows[months[2]][0][-1], "Early ship fr Oct'26 to Sep'26")
        self.assertEqual(rows[months[2]][1][0:2], (job, 850))
        self.assertAlmostEqual(rows[months[2]][1][2], 14350.09)
        self.assertEqual(rows[months[2]][1][-1], "")
        self.assertAlmostEqual(sum(row[2] for row in rows[months[2]]), 18028.72)

    def test_later_months_have_no_ordinary_missing_remark(self):
        months = ["2026-08", "2026-09", "2026-10"]
        bds = recon._new_month_maps(months)
        bds[months[2]]["AH001000LS"] = [100, 1000]
        rows = recon._build_reconciliation_rows(months[2], months, bds, recon._new_month_maps(months), {}, 0.5, 0.01)
        self.assertEqual(rows, [("AH001000LS", 100, 1000, 0, 0, "")])

    def test_previous_weekly_line_matches_partial_job_shipment(self):
        months = ["2026-08", "2026-09"]
        job = "AQ225025MJ"
        rows = {months[0]: [(job, 2, 32.48, 0, 0, "Missing")], months[1]: []}
        weekly = [
            recon.SourceLine(job, job, 1387, 22649.71, "", "", "July", 1),
            recon.SourceLine(job + "-TOP", job, 2, 20.12, "", "", "July", 2),
        ]
        recon._reconcile_missing_movements(rows, months, {"2026-07": {}}, recon._new_month_maps(months), {month: [] for month in months}, "2026-07", weekly, 0.5)
        self.assertEqual(rows[months[0]][0][-1], "Early ship fr Aug'26 to Jul'26")


if __name__ == "__main__":
    unittest.main()
