import importlib.util
import sys
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch


spec = importlib.util.spec_from_file_location("summary_reconcile", Path(__file__).with_name("main.py"))
recon = importlib.util.module_from_spec(spec)
sys.modules[spec.name] = recon
spec.loader.exec_module(recon)


class MovementReconciliationTests(unittest.TestCase):
    def test_weekly_descriptive_samples_and_amount_only_dn(self):
        wb = recon.Workbook()
        ws = wb.active
        ws.title = "Sep'26 Wk 3"
        for _ in range(6):
            ws.append([])
        for name, qty, planned, actual in [
            ("GGP SAMPLES", 805, 17109.05, 17109.05),
            ("GGP SAMPLES", 2116, 52416.04, 52416.04),
            ("GGP SAMPLES", 248, 5532.10, 5532.10),
            ("SAMPLES (MEETING)", 6, 156.96, 156.69),
            ("sample (Photoshoot )", 2, 62, 62),
            ("SAMPLES (SPA )", 61, 898, 898),
            ("SAMPLES (PHOTOSHOOT )", 42, 1096.83, 1096.83),
            ("SAMPLES (P1 )", 4, 49.20, 49.20),
            ("DN", None, 283, 283),
            ("SAMPLES", 100, 3342, 3342),
        ]:
            row = [None] * 18
            row[1], row[2], row[8] = name, qty, planned
            row[16], row[17] = qty, actual
            ws.append(row)
        with patch.object(recon, "load_workbook", return_value=wb):
            lines, sheets, unmatched, warnings = recon._read_weekly_actuals(
                "local.xlsx", "vn.xlsx", "2026-09", [], lambda message: None,
            )
        grouped = recon._aggregate_lines(lines)
        self.assertEqual(grouped["GGP SAMPLES"][0], 3169)
        self.assertAlmostEqual(grouped["GGP SAMPLES"][1], 75057.19)
        self.assertEqual(grouped["SAMPLES (MEETING)"], [6, 156.69])
        self.assertEqual(grouped["DN"], [0, 283])
        self.assertEqual(grouped["SAMPLES"], [100, 3342])
        self.assertEqual(len(grouped), 8)
        self.assertEqual(unmatched, [])
        self.assertEqual(warnings, [])
        months = ["2026-09", "2026-10"]
        bds = recon._new_month_maps(months)
        ann, _ = recon._build_ann_maps(
            months, {month: [] for month in months}, recon._new_month_maps(months),
            bds, lines, lambda message: None,
        )
        for job, pair in ann["2026-09"].items():
            reason = recon._classify_variance(job, months[0], months, bds, ann, {}, 0.5, 0.01)
            self.assertEqual(reason, "DN" if job == "DN" else "Salesman Sample")

    def test_bds_header_layouts_read_same_business_fields(self):
        for stacked in (False, True):
            with self.subTest(stacked=stacked):
                wb = recon.Workbook()
                ws = wb.active
                ws.title = "pcp2012"
                ws.append(["Date Prepared : 04/09/2026"])
                if stacked:
                    ws.append([None, None, None, None, "Order"])
                    ws.append([None, None, None, "GAC", "Qty", "Total"])
                    ws.append(["Job Number", "Job Number", "MCO", "Date", "FG", "Amount", "FtyLoc", "Jobtype"])
                else:
                    ws.append(
                        [
                            "Job Number",
                            "Job Number",
                            "MCO",
                            "GAC Date",
                            "Order Qty FG",
                            "Total Amount",
                            "FtyLoc",
                            "Jobtype",
                        ]
                    )
                ws.append(["WRONG", "AH001000LS", "VN", recon.date(2026, 9, 1), 5, 75, None, "B"])
                ws.append(["WRONG", "AH001001LS", "SIE_VN", recon.date(2026, 9, 1), 8, 80, None, "B"])
                ws.append(["WRONG", "AH001002LS", "VN", recon.date(2026, 9, 1), 9, 90, None, "Q"])
                ws.insert_cols(3, 2)
                with patch.object(recon, "load_workbook", return_value=wb):
                    bds, supplements, counts, months = recon._read_bds("bds.xlsx", lambda message: None)
                self.assertEqual(dict(bds["2026-09"]), {"AH001000LS": [5, 75]})
                self.assertEqual(dict(supplements["2026-09"]), {"AH001000LS": [5, 75]})
                self.assertEqual(counts["2026-09"], 1)
                self.assertEqual(sum(counts.values()), 1)
                self.assertEqual(months, recon._month_range("2026-09", "2027-04"))

    def test_bds_prepared_month_sets_reporting_range_and_previous_context(self):
        for prepared, start, end, previous in [
            ("04/12/2026", "2026-12", "2027-04", "2026-11"),
            ("04/01/2027", "2027-01", "2027-04", "2026-12"),
            ("04/04/2027", "2027-04", "2027-04", "2027-03"),
            ("04/05/2027", "2027-05", "2028-04", "2027-04"),
        ]:
            with self.subTest(prepared=prepared):
                wb = recon.Workbook()
                ws = wb.active
                ws.title = "pcp2012"
                ws.append([f"Date Prepared : {prepared}"])
                ws.append(["Job Number", "Jobtype", "MCO", "GAC Date", "Order Qty FG", "Total Amount", "Ftyloc"])
                ws.append(["AH001000LS", "B", "VN", recon.date.fromisoformat(previous + "-01"), 5, 75, "VN"])
                ws.append(["AH001001LS", "B", "VN", recon.date.fromisoformat(start + "-01"), 10, 100, "VN"])
                with patch.object(recon, "load_workbook", return_value=wb):
                    bds, supplements, counts, months = recon._read_bds("bds.xlsx", lambda message: None)
                self.assertEqual(months, recon._month_range(start, end))
                self.assertNotIn(previous, months)
                self.assertEqual(dict(bds[previous]), {"AH001000LS": [5, 75]})
                self.assertEqual(dict(bds[start]), {"AH001001LS": [10, 100]})
                self.assertEqual(counts[previous], 1)

    def test_bds_requires_valid_prepared_date(self):
        for prepared in (None, "Order Control List", "Date Prepared : 31/02/2026"):
            with self.subTest(prepared=prepared):
                wb = recon.Workbook()
                ws = wb.active
                ws.title = "pcp2012"
                ws.append([prepared])
                ws.append(["Job Number", "Jobtype", "MCO", "GAC Date", "Order Qty FG", "Total Amount", "Ftyloc"])
                with patch.object(recon, "load_workbook", return_value=wb):
                    with self.assertRaisesRegex(ValueError, "Date Prepared"):
                        recon._read_bds("bds.xlsx", lambda message: None)

    def test_bds_invalid_headers_fail_explicitly(self):
        for last_header, message in [("Wrong amount", "missing required column: amount"), ("Total Amount", "ambiguous columns for amount")]:
            with self.subTest(last_header=last_header):
                wb = recon.Workbook()
                ws = wb.active
                ws.append(["Job Number", "Jobtype", "MCO", "GAC Date", "Order Qty FG", last_header, last_header, "Ftyloc"])
                with self.assertRaisesRegex(ValueError, message):
                    recon._detect_bds_columns(ws)
        wb = recon.Workbook()
        wb.active.append(["Not an Order Control header"])
        with self.assertRaisesRegex(ValueError, "header row was not found"):
            recon._detect_bds_columns(wb.active)

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

    def test_forecast_preserves_reporting_range_despite_older_shipments(self):
        wb = recon.Workbook()
        ws = wb.active
        ws.title = "SHIPMENTS"
        ws.append(["JOB", "QTY", "AMOUNT", "PLAN EX-FTY"])
        ws.append(["AH001000LS", 10, 100, recon.date(2026, 8, 1)])
        ws.append(["AH001001LS", 20, 200, recon.date(2026, 12, 1)])
        ws.append(["AH001002LS", 30, 300, recon.date(2027, 5, 1)])
        months = recon._month_range("2026-11", "2027-04")
        with patch.object(recon, "load_workbook", return_value=wb):
            shipment, statuses, counts, warnings = recon._read_shipment_forecast(
                "forecast.xlsx", months, lambda message: None
            )
        self.assertEqual(list(shipment), months)
        self.assertEqual(counts["2026-11"], 0)
        self.assertEqual(counts["2027-04"], 0)
        self.assertEqual([line.full_job for line in shipment["2026-12"]], ["AH001001LS"])

    def test_reconciliation_uses_bds_month_for_forecast_actuals_and_movements(self):
        with TemporaryDirectory() as directory:
            paths = {role: str(Path(directory) / f"{role}.xlsx") for role in ("bsd", "shipment", "local", "vn")}
            wb = recon.Workbook()
            ws = wb.active
            ws.title = "pcp2012"
            ws.append(["Date Prepared : 04/11/2026"])
            ws.append(["Job Number", "Jobtype", "MCO", "GAC Date", "Order Qty FG", "Total Amount", "Ftyloc"])
            for job, month, qty, amount in [
                ("AH001000LS", 11, 10, 100),
                ("AH001001LS", 12, 20, 200),
                ("AH001002LS", 1, 30, 300),
                ("AH001003LS", 10, 7, 70),
                ("AH001004LS", 11, 4, 40),
                ("AH001005LS", 10, 8, 80),
                ("AH001006LS", 11, 9, 90),
            ]:
                ws.append([job, "B", "VN", recon.date(2027 if month == 1 else 2026, month, 1), qty, amount, "VN"])
            wb.save(paths["bsd"])
            wb.close()

            wb = recon.Workbook()
            ws = wb.active
            ws.title = "SHIPMENTS"
            ws.append(["JOB", "QTY", "AMOUNT", "PLAN EX-FTY", "INVOICE"])
            for job, month, qty, amount, invoice in [
                ("AH001007LS", 10, 100, 1000, "OCT-OLD"),
                ("AH001000LS", 11, 10, 100, "NOV-0"),
                ("AH001003LS", 11, 7, 77, "NOV-3"),
                ("AH001005LS", 11, 5, 55, "NOV-5"),
                ("AH001006LS", 11, 9, 90, "NOV-6"),
                ("AH001001LS", 12, 20, 210, "DEC-1"),
                ("AH001002LS", 1, 30, 300, "JAN-2"),
            ]:
                ws.append([job, qty, amount, recon.date(2027 if month == 1 else 2026, month, 1), invoice])
            wb.save(paths["shipment"])
            wb.close()

            for role in ("local", "vn"):
                wb = recon.Workbook()
                wb.remove(wb.active)
                for month in ("Oct", "Nov"):
                    ws = wb.create_sheet(f"{month}'26 Wk 1")
                    for _ in range(6):
                        ws.append([])
                    if role == "local":
                        shipments = (
                            [("AH001004LS", 4, 40, "OCT-4"), ("AH001005LS", 3, 30, "OCT-5")]
                            if month == "Oct"
                            else [("AH001006LS", 9, 99, "NOV-6")]
                        )
                        for job, qty, amount, invoice in shipments:
                            row = [None] * 18
                            row[1], row[2], row[3], row[8], row[16], row[17] = job, qty, invoice, amount, qty, amount
                            ws.append(row)
                wb.save(paths[role])
                wb.close()

            result = recon.process_reconciliation(paths, directory, recon.ProcessingOptions())
            self.assertEqual(result.reconciliation_months, recon._month_range("2026-11", "2027-04"))
            self.assertEqual(
                set(result.weekly_sheets),
                {f"{role}: {month}'26 Wk 1" for role in ("Local", "VN") for month in ("Oct", "Nov")},
            )
            wb = recon.load_workbook(result.output_path, read_only=True, data_only=False)
            try:
                self.assertNotIn("Recon Oct'26", wb.sheetnames)
                november = {row[0]: row for row in wb["Recon Nov'26"].iter_rows(min_row=4, values_only=True)}
                december = {row[0]: row for row in wb["Recon Dec'26"].iter_rows(min_row=4, values_only=True)}
                january = {row[0]: row for row in wb["Recon Jan'27"].iter_rows(min_row=4, values_only=True)}
                self.assertEqual(november["AH001000LS"][3:5], (10, 100))
                self.assertEqual(november["AH001006LS"][3:5], (9, 99))
                self.assertEqual(december["AH001001LS"][3:5], (20, 210))
                self.assertEqual(november["AH001003LS"][7], "Delay ship fr Oct'26 to Nov'26")
                self.assertEqual(november["AH001005LS"][7], "Delay ship fr Oct'26 to Nov'26")
                self.assertEqual(november["AH001004LS"][7], "Early ship fr Nov'26 to Oct'26")
                self.assertNotIn("AH001007LS", november)
                self.assertEqual(january["AH001002LS"][3:5], (None, None))
                self.assertIsNone(january["AH001002LS"][7])
                self.assertEqual(wb["Processing Log"]["A5"].value, "Reporting month (BDS Date Prepared)")
                self.assertEqual(wb["Processing Log"]["B5"].value, "Nov'26")
            finally:
                wb.close()

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
