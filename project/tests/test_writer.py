"""Regression tests for fixed-report numeric normalization and formula retention."""

import unittest

from openpyxl import Workbook
from openpyxl.comments import Comment

from engine.comments import CommentCopyStats
from engine.writer import EntryChangeStats, write_report_fixed


class WriteReportFixedTests(unittest.TestCase):
    def setUp(self):
        self.report_config = {
            "sheet_name": "固定报表",
            "data_start_row": 4,
            "data_start_col": "C",
            "data_end_col": "K",
            "formula_columns": ["H"],
            "text_columns": ["K"],
            "row_mapping": {4: "单位A"},
        }
        self.config = {
            "unit_to_owner": {"单位A": "权属A"},
            "ownership_files": {"权属A": {"file": "权属A.xlsx"}},
        }

    def _make_workbooks(self):
        template_book = Workbook()
        template_sheet = template_book.active
        template_sheet.title = "固定报表"
        template_sheet["B4"] = "单位A"
        for column in ("C", "D", "E", "F", "G", "I", "J", "K"):
            template_sheet[f"{column}4"] = "模板旧值"
        template_sheet["H4"] = "=F4/C4"

        source_book = Workbook()
        source_sheet = source_book.active
        source_sheet.title = "固定报表"
        source_sheet["B4"] = "单位A"
        return template_sheet, source_sheet, source_book

    def test_non_numeric_values_are_zeroed_and_template_formula_is_retained(self):
        template_sheet, source_sheet, source_book = self._make_workbooks()
        source_sheet["C4"] = None
        source_sheet["D4"] = 0
        source_sheet["E4"] = "/"
        source_sheet["F4"] = 12.5
        source_sheet["G4"] = None
        source_sheet["I4"] = "备注文本"
        source_sheet["J4"] = ""
        source_sheet["K4"] = "/"

        written_rows = write_report_fixed(
            template_sheet,
            self.report_config,
            {"权属A": {"workbook": source_book}},
            self.config,
            report_id=4,
        )

        self.assertEqual(written_rows, 1)
        self.assertEqual(template_sheet["C4"].value, 0)
        self.assertEqual(template_sheet["D4"].value, 0)
        self.assertEqual(template_sheet["E4"].value, 0)
        self.assertEqual(template_sheet["F4"].value, 12.5)
        self.assertEqual(template_sheet["G4"].value, 0)
        self.assertEqual(template_sheet["I4"].value, 0)
        self.assertEqual(template_sheet["J4"].value, 0)
        self.assertEqual(template_sheet["K4"].value, "/")
        self.assertEqual(template_sheet["H4"].value, "=F4/C4")

    def test_blank_source_row_is_zeroed_and_keeps_template_formula(self):
        template_sheet, _, source_book = self._make_workbooks()

        written_rows = write_report_fixed(
            template_sheet,
            self.report_config,
            {"权属A": {"workbook": source_book}},
            self.config,
            report_id=4,
        )

        self.assertEqual(written_rows, 1)
        for column in ("C", "D", "E", "F", "G", "I", "J"):
            self.assertEqual(template_sheet[f"{column}4"].value, 0)
        self.assertIsNone(template_sheet["K4"].value)
        self.assertEqual(template_sheet["H4"].value, "=F4/C4")

    def test_excluded_owner_is_not_written(self):
        template_sheet, source_sheet, source_book = self._make_workbooks()
        source_sheet["C4"] = 99

        written_rows = write_report_fixed(
            template_sheet,
            self.report_config,
            {"权属A": {"workbook": source_book}},
            self.config,
            report_id=4,
            excluded_owners={"权属A"},
        )

        self.assertEqual(written_rows, 0)
        self.assertEqual(template_sheet["C4"].value, "模板旧值")

    def test_unconfigured_source_unit_is_warned_and_not_written(self):
        template_sheet, source_sheet, source_book = self._make_workbooks()
        source_sheet["C4"] = 10
        source_sheet["A5"] = "权属A"
        source_sheet["B5"] = "新增单位"
        source_sheet["C5"] = 999
        source_sheet["A6"] = "其他权属"
        source_sheet["B6"] = "其他权属的单位"
        source_sheet["C6"] = 888
        source_sheet["A7"] = "合计"
        source_sheet["B8"] = "合计行后的说明文字"

        entry_changes = EntryChangeStats()
        with self.assertLogs("engine.writer", level="WARNING") as captured:
            written_rows = write_report_fixed(
                template_sheet,
                self.report_config,
                {"权属A": {"workbook": source_book}},
                self.config,
                report_id=1,
                entry_change_stats=entry_changes,
            )

        warning_text = "\n".join(captured.output)
        self.assertEqual(written_rows, 1)
        self.assertIn("[未配置单位检测]", warning_text)
        self.assertIn("权属：权属A", warning_text)
        self.assertIn("'新增单位'（源表第5行）", warning_text)
        self.assertIn("处理结果：未写入汇总表", warning_text)
        self.assertNotIn("其他权属的单位", warning_text)
        self.assertNotIn("合计行后的说明文字", warning_text)
        self.assertEqual(template_sheet["C4"].value, 10)
        self.assertIsNone(template_sheet["B5"].value)
        self.assertIsNone(template_sheet["C5"].value)
        self.assertEqual(entry_changes.added, 1)
        self.assertEqual(entry_changes.removed, 0)
        self.assertEqual(entry_changes.details[0].unit_name, "新增单位")
        self.assertEqual(entry_changes.details[0].source_row, 5)

    def test_missing_configured_unit_is_recorded_as_removed(self):
        template_sheet, source_sheet, source_book = self._make_workbooks()
        source_sheet["B4"] = "其他单位"
        entry_changes = EntryChangeStats()

        with self.assertLogs("engine.writer", level="WARNING"):
            written_rows = write_report_fixed(
                template_sheet,
                self.report_config,
                {"权属A": {"workbook": source_book}},
                self.config,
                report_id=1,
                entry_change_stats=entry_changes,
            )

        self.assertEqual(written_rows, 0)
        self.assertEqual(entry_changes.added, 0)
        self.assertEqual(entry_changes.removed, 1)
        self.assertEqual(entry_changes.details[0].unit_name, "单位A")
        self.assertEqual(entry_changes.details[0].target_row, 4)

    def test_source_comment_is_copied_with_author_and_logged(self):
        template_sheet, source_sheet, source_book = self._make_workbooks()
        source_sheet["C4"] = 245923.54
        source_sheet["C4"].comment = Comment(
            "已剔除非自有土地面积", "Administrator"
        )
        stats = CommentCopyStats()

        with self.assertLogs("engine.comments", level="DEBUG") as captured:
            write_report_fixed(
                template_sheet,
                self.report_config,
                {"权属A": {"workbook": source_book}},
                self.config,
                report_id=1,
                comment_stats=stats,
            )

        target_comment = template_sheet["C4"].comment
        self.assertIsNotNone(target_comment)
        self.assertEqual(target_comment.text, "已剔除非自有土地面积")
        self.assertEqual(target_comment.author, "Administrator")
        self.assertEqual(source_sheet["C4"].comment.author, "Administrator")
        self.assertEqual(stats.copied, 1)
        self.assertEqual(len(stats.details), 1)
        detail = stats.details[0]
        self.assertEqual(detail.report_id, 1)
        self.assertEqual(detail.owner_key, "权属A")
        self.assertEqual(detail.source_sheet, "固定报表")
        self.assertEqual(detail.source_cell, "C4")
        self.assertEqual(detail.target_sheet, "固定报表")
        self.assertEqual(detail.target_cell, "C4")
        self.assertEqual(detail.author, "Administrator")
        self.assertEqual(detail.text, "已剔除非自有土地面积")
        log_text = "\n".join(captured.output)
        self.assertIn("批注已复制", log_text)
        self.assertIn("固定报表!C4", log_text)
        self.assertIn("作者=Administrator", log_text)
        self.assertIn("已剔除非自有土地面积", log_text)

    def test_configured_difference_is_warned_before_source_overwrites_template(self):
        template_sheet, source_sheet, source_book = self._make_workbooks()
        report_config = dict(self.report_config)
        report_config["difference_check_columns"] = ["C", "E"]
        template_sheet["C3"] = "环比基准字段"
        template_sheet["E3"] = "同比基准字段"
        template_sheet["C4"] = 100
        template_sheet["D4"] = 200
        template_sheet["E4"] = 300
        source_sheet["C4"] = 110
        source_sheet["D4"] = 999
        source_sheet["E4"] = 300

        with self.assertLogs("engine.writer", level="WARNING") as captured:
            written_rows = write_report_fixed(
                template_sheet,
                report_config,
                {
                    "权属A": {
                        "workbook": source_book,
                        "filename": "权属A原始表.xlsx",
                    }
                },
                self.config,
                report_id=3,
            )

        warning_text = "\n".join(captured.output)
        self.assertEqual(written_rows, 1)
        self.assertIn("[环比/同比数据差异检测]", warning_text)
        self.assertIn("报表3 固定报表", warning_text)
        self.assertIn("权属：权属A", warning_text)
        self.assertIn("单位：单位A", warning_text)
        self.assertIn("字段：环比基准字段（C列）", warning_text)
        self.assertIn("模板原值：固定报表!C4=100", warning_text)
        self.assertIn("权属文件值：权属A原始表.xlsx/固定报表!C4=110", warning_text)
        self.assertNotIn("D列", warning_text)
        self.assertNotIn("E列", warning_text)
        self.assertEqual(template_sheet["C4"].value, 110)

    def test_equal_configured_values_do_not_emit_difference_warning(self):
        template_sheet, source_sheet, source_book = self._make_workbooks()
        report_config = dict(self.report_config)
        report_config["difference_check_columns"] = ["C", "E"]
        template_sheet["C4"] = 100
        template_sheet["E4"] = 300
        source_sheet["C4"] = 100
        source_sheet["E4"] = 300

        with self.assertNoLogs("engine.writer", level="WARNING"):
            written_rows = write_report_fixed(
                template_sheet,
                report_config,
                {"权属A": {"workbook": source_book}},
                self.config,
                report_id=1,
            )

        self.assertEqual(written_rows, 1)


if __name__ == "__main__":
    unittest.main()
