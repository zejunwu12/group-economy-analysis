"""日志展示层回归测试。"""

import logging
import tempfile
import unittest

from engine.comments import CommentCopyDetail, CommentCopyStats
from logger import (
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
        self.assertIn("[报表处理结果]", log_text)
        self.assertIn("[数据完整性和异常值检测]", log_text)
        self.assertIn("[批注处理]", log_text)
        self.assertIn("[生成文件]", log_text)

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
