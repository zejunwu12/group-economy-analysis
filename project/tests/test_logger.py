"""日志展示层回归测试。"""

import logging
import io
import tempfile
import unittest

from engine.comments import CommentCopyDetail, CommentCopyStats
from logger import (
    get_suppressed_console_warning_count,
    log_ownership_check_details,
    log_processing_summary,
    log_report_step,
    log_workflow_step,
    setup_logger,
)


class LoggerSetupTests(unittest.TestCase):
    def tearDown(self):
        root = logging.getLogger()
        for handler in root.handlers[:]:
            root.removeHandler(handler)
            handler.close()

    def test_console_is_concise_and_file_keeps_debug_details(self):
        with tempfile.TemporaryDirectory() as output_dir:
            try:
                log_path = setup_logger("INFO", output_dir)

                handlers = {
                    handler.get_name(): handler
                    for handler in logging.getLogger().handlers
                }
                self.assertEqual(handlers["summary-console"].level, logging.INFO)
                self.assertEqual(handlers["summary-file"].level, logging.DEBUG)

                logging.getLogger("test.detail").debug("逐项追溯信息")
                handlers["summary-file"].flush()
                with open(log_path, encoding="utf-8") as log_file:
                    file_text = log_file.read()

                self.assertIn("逐项追溯信息", file_text)
            finally:
                self.tearDown()

    def test_console_hides_unmarked_details_and_file_retains_them(self):
        with tempfile.TemporaryDirectory() as output_dir:
            try:
                log_path = setup_logger("INFO", output_dir)
                handlers = {
                    handler.get_name(): handler
                    for handler in logging.getLogger().handlers
                }
                console_stream = io.StringIO()
                handlers["summary-console"].setStream(console_stream)
                detail_logger = logging.getLogger("test.concise")

                detail_logger.info("普通处理明细")
                detail_logger.warning("普通警告明细")
                detail_logger.info(
                    "控制台摘要",
                    extra={"console_summary": True},
                )
                handlers["summary-file"].flush()

                console_text = console_stream.getvalue()
                with open(log_path, encoding="utf-8") as log_file:
                    file_text = log_file.read()

                self.assertEqual(get_suppressed_console_warning_count(), 1)
                self.assertIn("控制台摘要", console_text)
                self.assertNotIn("普通处理明细", console_text)
                self.assertNotIn("普通警告明细", console_text)
                self.assertIn("普通处理明细", file_text)
                self.assertIn("普通警告明细", file_text)
                self.assertIn("控制台摘要", file_text)
            finally:
                self.tearDown()

    def test_workflow_and_report_steps_have_clear_hierarchy(self):
        workflow_logger = logging.getLogger("test.workflow")

        with self.assertLogs("test.workflow", level="INFO") as captured:
            log_workflow_step(4, 6, "写入各报表", step_logger=workflow_logger)
            log_report_step(
                1,
                2,
                1,
                "报表1 资产总体情况表",
                step_logger=workflow_logger,
            )

        log_text = "\n".join(captured.output)
        self.assertIn("[步骤 4/6] 写入各报表", log_text)
        self.assertIn("[报表 1/2] 报表1 资产总体情况表", log_text)
        self.assertNotIn("报表1 报表1", log_text)

    def test_processing_summary_uses_action_specific_labels(self):
        summary_logger = logging.getLogger("test.summary")
        config = {
            "quarter": {"label": "测试周期"},
            "ownership_files": {
                "权属A": {"file": "权属A.xlsx"},
                "集团本部": {"file": None},
            },
            "runtime": {"reports_to_run": [1]},
            "reports": {"report1": {"sheet_name": "报表1 测试表"}},
        }

        with self.assertLogs("test.summary", level="INFO") as captured:
            log_processing_summary(
                config=config,
                ownership_data={},
                report_results={1: 1},
                summary_logger=summary_logger,
            )

        log_text = "\n".join(captured.output)
        self.assertIn("运行结果摘要", log_text)
        self.assertIn(
            "[权属文件加载] 已加载 0 个，未加载 1 个，无需加载 1 个",
            log_text,
        )
        self.assertIn("[权属表格式检查]", log_text)
        self.assertIn(
            "  [权属表格式检查] 通过 0 项，需核对 0 项",
            log_text,
        )
        self.assertIn("[报表处理结果]", log_text)
        self.assertIn("[数据完整性和异常值检测]", log_text)
        self.assertIn("[批注处理]", log_text)
        self.assertIn("[生成文件]", log_text)

    def test_processing_summary_lists_header_mismatches_after_file_loading(self):
        summary_logger = logging.getLogger("test.summary.headers")
        config = {
            "quarter": {"label": "测试周期"},
            "ownership_files": {"古城集团": {"file": "古城集团.xlsx"}},
            "runtime": {"reports_to_run": [8]},
            "reports": {"report8": {"sheet_name": "报表8 房屋土地明细表"}},
        }
        header_mismatches = {
            ("古城集团", 8): "表头有效列范围不一致（模板 A:L，权属表 A:M）"
        }

        with self.assertLogs("test.summary.headers", level="DEBUG") as captured:
            summary = log_processing_summary(
                config=config,
                ownership_data={"古城集团": {"filename": "古城集团.xlsx"}},
                report_results={8: {"record_count": 0, "data_end_row": 20}},
                header_mismatches=header_mismatches,
                summary_logger=summary_logger,
            )

        log_text = "\n".join(captured.output)
        self.assertLess(
            log_text.index("[权属文件加载]"),
            log_text.index("[权属表格式检查]"),
        )
        self.assertLess(
            log_text.index("[权属表格式检查]"),
            log_text.index("[报表处理结果]"),
        )
        self.assertIn(
            "  [权属表格式检查] 通过 0 项，需核对 1 项",
            log_text,
        )
        self.assertIn("    需核对问题明细：", log_text)
        self.assertIn(
            "[问题1] 报表8 房屋土地明细表｜权属：古城集团",
            log_text,
        )
        self.assertIn(
            "原因：表头有效列范围不一致（模板 A:L，权属表 A:M）",
            log_text,
        )
        self.assertIn(
            "处理结果：该权属本报表未写入，请人工复核",
            log_text,
        )
        self.assertEqual(summary["header_mismatches"], 1)

    def test_header_mismatch_details_are_written_to_file_not_console(self):
        config = {
            "quarter": {"label": "测试周期"},
            "ownership_files": {"古城集团": {"file": "古城集团.xlsx"}},
            "runtime": {"reports_to_run": [8]},
            "reports": {"report8": {"sheet_name": "报表8 房屋土地明细表"}},
        }
        header_mismatches = {
            ("古城集团", 8): "表头有效列范围不一致（模板 A:L，权属表 A:M）"
        }

        with tempfile.TemporaryDirectory() as output_dir:
            try:
                log_path = setup_logger("INFO", output_dir)
                handlers = {
                    handler.get_name(): handler
                    for handler in logging.getLogger().handlers
                }
                console_stream = io.StringIO()
                handlers["summary-console"].setStream(console_stream)

                log_processing_summary(
                    config=config,
                    ownership_data={"古城集团": {"filename": "古城集团.xlsx"}},
                    report_results={8: {"record_count": 0, "data_end_row": 20}},
                    header_mismatches=header_mismatches,
                )

                console_text = console_stream.getvalue()
                with open(log_path, encoding="utf-8") as log_file:
                    file_text = log_file.read()

                self.assertIn(
                    "[权属表格式检查] 通过 0 项，需核对 1 项",
                    console_text,
                )
                self.assertNotIn("需核对问题明细", console_text)
                self.assertNotIn("权属：古城集团", console_text)
                self.assertIn("需核对问题明细", file_text)
                self.assertIn("权属：古城集团", file_text)
                self.assertIn(
                    "处理结果：该权属本报表未写入，请人工复核",
                    file_text,
                )
            finally:
                self.tearDown()

    def test_step_three_ownership_details_match_summary_and_stay_off_console(self):
        config = {
            "quarter": {"label": "测试周期"},
            "ownership_files": {
                "古城集团": {"file": "古城集团.xlsx"},
                "集团本部": {"file": "集团本部.xlsx"},
            },
            "runtime": {"reports_to_run": [8]},
            "reports": {"report8": {"sheet_name": "报表8 房屋土地明细表"}},
        }
        ownership_data = {"古城集团": {"filename": "古城集团.xlsx"}}
        header_mismatches = {
            ("古城集团", 8): "表头有效列范围不一致（模板 A:L，权属表 A:M）"
        }

        with tempfile.TemporaryDirectory() as output_dir:
            try:
                log_path = setup_logger("INFO", output_dir)
                handlers = {
                    handler.get_name(): handler
                    for handler in logging.getLogger().handlers
                }
                console_stream = io.StringIO()
                handlers["summary-console"].setStream(console_stream)

                log_workflow_step(3, 6, "加载并检查权属数据文件")
                log_ownership_check_details(
                    config,
                    ownership_data,
                    header_mismatches,
                )
                handlers["summary-file"].flush()

                console_text = console_stream.getvalue()
                with open(log_path, encoding="utf-8") as log_file:
                    file_text = log_file.read()

                step_position = file_text.index(
                    "[步骤 3/6] 加载并检查权属数据文件"
                )
                loading_position = file_text.index(
                    "[权属文件加载] 已加载 1 个，未加载 1 个，无需加载 0 个"
                )
                mismatch_position = file_text.index(
                    "[权属表格式检查] 通过 0 项，需核对 1 项"
                )
                self.assertLess(step_position, loading_position)
                self.assertLess(loading_position, mismatch_position)
                self.assertIn("集团本部: 未加载", file_text)
                self.assertIn("权属：古城集团", file_text)
                self.assertIn(
                    "处理结果：该权属本报表未写入，请人工复核",
                    file_text,
                )
                self.assertIn("[步骤 3/6] 加载并检查权属数据文件", console_text)
                self.assertNotIn("[权属文件加载]", console_text)
                self.assertNotIn("权属：古城集团", console_text)
            finally:
                self.tearDown()

    def test_processing_summary_lists_validation_problem_details(self):
        summary_logger = logging.getLogger("test.summary.details")
        config = {
            "quarter": {"label": "测试周期"},
            "ownership_files": {},
            "runtime": {"reports_to_run": [4]},
            "reports": {"report4": {"sheet_name": "报表4 各企业出租率"}},
        }
        total_validation = {
            "passed": 1,
            "warnings": 1,
            "skipped": 0,
            "details": [
                {
                    "report_id": 4,
                    "sheet_name": "报表4 各企业出租率",
                    "column": "G",
                    "status": "warning",
                    "formula": "=SUM(G4:G12)",
                    "formula_range": "G4:G12",
                    "formula_sum": 100,
                    "expected_sum": 120,
                    "difference": -20,
                    "reason": "SUM 公式范围计算结果与完整数据区合计不一致",
                }
            ],
        }
        anomaly_report = {
            "total": 1,
            "negative_area": 0,
            "oversized_area": 0,
            "empty_key_data": 1,
            "details": [
                {
                    "report_id": 2,
                    "sheet_name": "报表2 按建筑类型区分总资产",
                    "cell": "B4",
                    "message": "关键字段均为空",
                    "value": "集团本部",
                }
            ],
        }

        with self.assertLogs("test.summary.details", level="INFO") as captured:
            log_processing_summary(
                config=config,
                ownership_data={},
                report_results={4: 1},
                total_validation=total_validation,
                anomaly_report=anomaly_report,
                summary_logger=summary_logger,
            )

        log_text = "\n".join(captured.output)
        self.assertIn("需核对问题明细", log_text)
        self.assertIn("[问题1] 报表4 各企业出租率｜G列", log_text)
        self.assertIn("合计公式：=SUM(G4:G12)｜引用范围：G4:G12", log_text)
        self.assertIn(
            "公式范围合计：100｜完整数据区合计：120｜"
            "差额（公式范围合计-完整数据区合计）：-20",
            log_text,
        )
        self.assertIn("异常明细", log_text)
        self.assertIn(
            "[问题1] 报表2 按建筑类型区分总资产｜位置：B4｜"
            "关键字段均为空｜标识值：集团本部",
            log_text,
        )

    def test_processing_summary_lists_copied_comment_details(self):
        summary_logger = logging.getLogger("test.summary.comments")
        config = {
            "quarter": {"label": "测试周期"},
            "ownership_files": {},
            "runtime": {"reports_to_run": [3]},
            "reports": {"report3": {"sheet_name": "报表3 资产运营情况表"}},
        }
        comment_stats = CommentCopyStats(
            copied=1,
            details=[
                CommentCopyDetail(
                    report_id=3,
                    owner_key="商业集团",
                    source_sheet="报表3 资产运营情况表",
                    source_cell="J5",
                    target_sheet="报表3 资产运营情况表",
                    target_cell="J8",
                    author="填报人",
                    text="第一行\n第二行",
                )
            ],
        )

        with self.assertLogs("test.summary.comments", level="INFO") as captured:
            log_processing_summary(
                config=config,
                ownership_data={},
                report_results={3: 1},
                comment_stats=comment_stats,
                summary_logger=summary_logger,
            )

        log_text = "\n".join(captured.output)
        self.assertIn("权属批注复制明细", log_text)
        self.assertIn(
            "[批注1] 报表3 资产运营情况表｜权属：商业集团",
            log_text,
        )
        self.assertIn(
            "来源：报表3 资产运营情况表!J5｜"
            "目标：报表3 资产运营情况表!J8｜作者：填报人",
            log_text,
        )
        self.assertIn("内容：第一行 / 第二行", log_text)


if __name__ == "__main__":
    unittest.main()
