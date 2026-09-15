import unittest
from collections import OrderedDict

from main import DivertEvent, DivertFromEvent, POGroup, build_allocations_from_divert_from_records, build_target_results


class DivertAllocationTests(unittest.TestCase):
    def test_equal_target_totals_respect_source_sizes_in_either_order(self):
        records = [
            DivertFromEvent(2, 100, 3, 5, 0, "00103"),
            DivertFromEvent(1, 200, 3, 7, 1, "00203"),
            DivertFromEvent(2, 100, 1, 12, 2, "00101"),
        ]
        limits = {(1, 200): {"2XL": 7}, (2, 100): {"XS": 12, "2XL": 5}}
        for sizes in [[("XS", 12), ("2XL", 12)], [("2XL", 12), ("XS", 12)]]:
            with self.subTest(sizes=sizes):
                result = build_allocations_from_divert_from_records(
                    POGroup(2, 400, OrderedDict(sizes)), records, limits
                )
                self.assertEqual(result.allocations, limits)
                self.assertEqual(result.residuals, {})

    def test_direct_references_do_not_depend_on_target_row_order(self):
        original = {(1, 100): OrderedDict([("S", 54), ("L", 29)])}
        events = [
            DivertEvent(1, 100, 2, 100, 1, 54, 1, "00101"),
            DivertEvent(1, 100, 2, 100, 2, 29, 1, "00102"),
        ]
        for sizes in [[("S", 54), ("L", 29)], [("L", 29), ("S", 54)]]:
            with self.subTest(sizes=sizes):
                groups = {
                    (1, 100): POGroup(1, 100, OrderedDict([("S", 0), ("L", 0)])),
                    (2, 100): POGroup(2, 100, OrderedDict(sizes)),
                }
                result = build_target_results(groups, original, events)[0]
                self.assertEqual(result.allocations, {(1, 100): {"S": 54, "L": 29}})
                self.assertEqual(result.residuals, {})

    def test_combined_records_cannot_exceed_source_size_quantity(self):
        records = [DivertFromEvent(1, 100, 1, 5, i, "00101") for i in range(2)]
        result = build_allocations_from_divert_from_records(
            POGroup(2, 100, OrderedDict([("XS", 10)])), records, {(1, 100): {"XS": 5}}
        )
        self.assertIsNone(result)

    def test_equal_quantities_from_different_sources_remain_candidates(self):
        records = [
            DivertFromEvent(1, 100, 1, 5, 0, "00101"),
            DivertFromEvent(2, 100, 1, 5, 1, "00101"),
            DivertFromEvent(1, 100, 2, 10, 2, "00102"),
        ]
        result = build_allocations_from_divert_from_records(
            POGroup(3, 100, OrderedDict([("XS", 5), ("S", 15)])),
            records,
            {(1, 100): {"XS": 5, "S": 15}, (2, 100): {"XS": 5}},
        )
        self.assertEqual(result.allocations, {(2, 100): {"XS": 5}, (1, 100): {"S": 15}})


if __name__ == "__main__":
    unittest.main()
