"""Quarter runtime parameter regression tests."""

import unittest

from engine.period import QuarterContext, QuarterError
from main import _parse_args


class QuarterContextTests(unittest.TestCase):
    def test_parses_quarter_and_derives_chinese_metadata(self):
        period = QuarterContext.parse("2026Q2")

        self.assertEqual(period.code, "2026Q2")
        self.assertEqual(period.label, "2026年第二季度")
        self.assertEqual(period.start_date, "2026年1月1日")
        self.assertEqual(period.end_date, "2026年6月30日")
        self.assertEqual(
            period.as_config(),
            {
                "code": "2026Q2",
                "year": 2026,
                "quarter": 2,
                "label": "2026年第二季度",
                "start_date": "2026年1月1日",
                "end_date": "2026年6月30日",
            },
        )

    def test_normalizes_lowercase_quarter_code(self):
        self.assertEqual(QuarterContext.parse(" 2026q3 ").code, "2026Q3")

    def test_next_quarter_crosses_year_boundary(self):
        next_period = QuarterContext.parse("2026Q4").next()

        self.assertEqual(next_period.code, "2027Q1")
        self.assertEqual(next_period.label, "2027年第一季度")
        self.assertEqual(next_period.end_date, "2027年3月31日")

    def test_rejects_invalid_quarter_code(self):
        for invalid in ("2026", "2026Q0", "2026Q5", "26Q2", "第二季度"):
            with self.subTest(invalid=invalid):
                with self.assertRaises(QuarterError):
                    QuarterContext.parse(invalid)

    def test_command_line_requires_quarter(self):
        with self.assertRaises(SystemExit) as captured:
            _parse_args([])

        self.assertEqual(captured.exception.code, 2)

    def test_command_line_accepts_quarter_and_optional_config(self):
        args = _parse_args(
            ["--quarter", "2026q2", "--config", "alternate.yaml"]
        )

        self.assertEqual(args.quarter, "2026Q2")
        self.assertEqual(str(args.config), "alternate.yaml")

        alias_args = _parse_args(
            ["--quarter", "2026Q2", "--temp", "final.xlsx"]
        )
        self.assertEqual(str(alias_args.template), "final.xlsx")


if __name__ == "__main__":
    unittest.main()
