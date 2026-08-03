"""Main entry point for the group economic analysis workflow."""

import argparse
import logging
import sys
from io import BytesIO
from pathlib import Path

import openpyxl

from engine.config_loader import ConfigLoader
from engine.comments import CommentCopyStats, clear_template_comments
from engine.output import save_summary_workbook
from engine.period import QuarterContext, QuarterError
from engine.reader import (
    close_ownership_files,
    load_ownership_files,
    load_template,
    validate_template,
)
from engine.report5_handler import process_report5
from engine.report7_handler import process_report7
from engine.report8_handler import process_report8
from engine.validator import detect_anomalies, validate_totals
from engine.writer import write_report_fixed
from logger import (
    log_processing_summary,
    log_report_step,
    log_workflow_step,
    setup_logger,
)


logger = logging.getLogger(__name__)
_FIXED_REPORT_IDS = frozenset({1, 2, 3, 4, 6})
_WORKFLOW_STEP_COUNT = 6


def run(
    quarter_code: str,
    config_path: str | Path | None = None,
    template_path: str | Path | None = None,
) -> dict:
    """Execute the complete eight-report summary workflow.

    Args:
        quarter_code: Reporting quarter in strict ``YYYYQn`` format.
        config_path: Optional path to a reusable static YAML configuration.
        template_path: Optional runtime template override. When omitted,
            ``runtime.template_path`` from the selected config is used.

    Returns a dictionary containing the output path, log path, per-report
    results, total validation report and anomaly report.
    """
    quarter = QuarterContext.parse(quarter_code)
    resolved_config_path = _resolve_config_path(config_path)
    config = ConfigLoader(str(resolved_config_path)).load()
    # Keep the existing downstream interface while making the command-line
    # argument the only persisted source of reporting-period information.
    config["quarter"] = quarter.as_config()
    output_dir = _resolve_runtime_path(config, config["runtime"]["output_dir"])
    log_path = setup_logger(config["runtime"]["log_level"], str(output_dir))

    workbook = None
    ownership_data = {}
    report_results = {}
    report_errors = {}
    validation_errors = {}
    total_validation = None
    anomaly_report = None
    output_path = None
    comment_stats = CommentCopyStats()

    log_workflow_step(1, _WORKFLOW_STEP_COUNT, "运行准备", step_logger=logger)
    logger.info(f"汇总周期: {config['quarter']['label']}")
    logger.info(f"配置文件: {resolved_config_path}")
    reports_to_run = tuple(config["runtime"].get("reports_to_run", range(1, 9)))
    logger.info(
        "计划处理报表: " + "、".join(str(report_id) for report_id in reports_to_run)
    )

    try:
        log_workflow_step(
            2,
            _WORKFLOW_STEP_COUNT,
            "加载并校验汇总模板",
            step_logger=logger,
        )
        workbook = load_template(config, template_path=template_path)
        validate_template(workbook, config, report_ids=reports_to_run)
        clear_template_comments(workbook, stats=comment_stats)

        log_workflow_step(
            3,
            _WORKFLOW_STEP_COUNT,
            "加载权属数据文件",
            step_logger=logger,
        )
        ownership_data = load_ownership_files(config)

        log_workflow_step(
            4,
            _WORKFLOW_STEP_COUNT,
            "写入各报表",
            step_logger=logger,
        )
        for report_position, report_id in enumerate(reports_to_run, start=1):
            report_config = config["reports"][f"report{report_id}"]
            log_report_step(
                report_position,
                len(reports_to_run),
                report_id,
                report_config["sheet_name"],
                step_logger=logger,
            )
            snapshot = _snapshot_workbook(workbook)
            comments_checkpoint = comment_stats.checkpoint()
            try:
                report_results[report_id] = _process_report(
                    report_id,
                    workbook,
                    ownership_data,
                    config,
                    comment_stats,
                )
            except Exception as exc:
                report_errors[report_id] = f"{type(exc).__name__}: {exc}"
                logger.exception(
                    f"报表{report_id}处理失败，已跳过并继续后续报表"
                )
                workbook.close()
                workbook = _restore_workbook(snapshot)
                rolled_back_comments = comment_stats.rollback(comments_checkpoint)
                logger.warning(f"报表{report_id}已恢复到处理前状态")
                if rolled_back_comments:
                    logger.warning(
                        f"报表{report_id}已撤销 {rolled_back_comments} 条批注复制记录"
                    )
            finally:
                snapshot.close()

        log_workflow_step(
            5,
            _WORKFLOW_STEP_COUNT,
            "执行数据校验",
            step_logger=logger,
        )
        if config["runtime"].get("enable_validation", True):
            fixed_report_ids = tuple(
                report_id
                for report_id in report_results
                if report_id in _FIXED_REPORT_IDS
            )
            if fixed_report_ids:
                try:
                    total_validation = validate_totals(
                        workbook,
                        config,
                        report_ids=fixed_report_ids,
                    )
                except Exception as exc:
                    validation_errors["totals"] = (
                        f"{type(exc).__name__}: {exc}"
                    )
                    logger.exception("合计值校验失败，已跳过并继续保存")

            successful_report_ids = tuple(report_results)
            if successful_report_ids:
                try:
                    anomaly_report = detect_anomalies(
                        workbook,
                        config,
                        report_ids=successful_report_ids,
                    )
                except Exception as exc:
                    validation_errors["anomalies"] = (
                        f"{type(exc).__name__}: {exc}"
                    )
                    logger.exception("数据完整性和异常值检测失败，已跳过并继续保存")
        else:
            logger.warning(
                "数据校验未执行（配置项 runtime.enable_validation=false）"
            )

        log_workflow_step(
            6,
            _WORKFLOW_STEP_COUNT,
            "保存文件并输出摘要",
            step_logger=logger,
        )
        output_path = save_summary_workbook(workbook, config)
        summary = log_processing_summary(
            config=config,
            ownership_data=ownership_data,
            report_results=report_results,
            total_validation=total_validation,
            anomaly_report=anomaly_report,
            output_path=output_path,
            report_errors=report_errors,
            comment_stats=comment_stats,
        )
        missing_owners = sorted(
            owner_key
            for owner_key, entry in config["ownership_files"].items()
            if entry.get("file") is not None and owner_key not in ownership_data
        )
        if report_errors or validation_errors or missing_owners:
            completion_notes = []
            if missing_owners:
                completion_notes.append(
                    "未加载权属文件：" + "、".join(missing_owners)
                )
            if report_errors:
                completion_notes.append(
                    "处理失败报表："
                    + "、".join(str(report_id) for report_id in sorted(report_errors))
                )
            if validation_errors:
                completion_notes.append(
                    "执行失败校验：" + "、".join(sorted(validation_errors))
                )
            logger.warning(
                "汇总文件已生成；" + "；".join(completion_notes)
            )
        else:
            logger.info("汇总文件已生成，处理流程完成")
        return {
            "output_path": output_path,
            "log_path": log_path,
            "report_results": report_results,
            "report_errors": report_errors,
            "validation_errors": validation_errors,
            "total_validation": total_validation,
            "anomaly_report": anomaly_report,
            "summary": summary,
        }
    finally:
        close_ownership_files(ownership_data)
        if workbook is not None:
            workbook.close()


def _process_report(
    report_id: int,
    workbook,
    ownership_data: dict,
    config: dict,
    comment_stats: CommentCopyStats,
):
    """Dispatch one report to its corresponding writer."""
    report_config = config["reports"][f"report{report_id}"]
    worksheet = workbook[report_config["sheet_name"]]
    logger.debug(f"开始处理报表{report_id}: {worksheet.title}")

    if report_id in _FIXED_REPORT_IDS:
        result = write_report_fixed(
            worksheet,
            report_config,
            ownership_data,
            config,
            report_id,
            comment_stats,
        )
    elif report_id == 5:
        result = process_report5(
            worksheet, ownership_data, report_config, comment_stats,
        )
    elif report_id == 7:
        result = process_report7(
            worksheet, ownership_data, report_config, comment_stats,
        )
    elif report_id == 8:
        result = process_report8(
            worksheet, ownership_data, report_config, comment_stats,
        )
    else:
        raise ValueError(f"不支持的报表编号: {report_id}")

    logger.debug(f"报表{report_id}处理完成")
    return result


def _resolve_config_path(config_path: str | Path | None) -> Path:
    if config_path is None:
        return Path(__file__).resolve().with_name("config.yaml")
    return Path(config_path).expanduser().resolve()


def _snapshot_workbook(workbook) -> BytesIO:
    """Create an in-memory checkpoint before processing one report."""
    snapshot = BytesIO()
    workbook.save(snapshot)
    snapshot.seek(0)
    return snapshot


def _restore_workbook(snapshot: BytesIO):
    """Restore a workbook checkpoint after a report-level failure."""
    snapshot.seek(0)
    return openpyxl.load_workbook(
        snapshot,
        read_only=False,
        data_only=False,
        keep_links=True,
    )


def _resolve_runtime_path(config: dict, configured_path: str) -> Path:
    config_dir = Path(config["runtime"]["_config_dir"])
    return (config_dir / configured_path).resolve()


def _quarter_argument(value: str) -> str:
    """Validate an argparse quarter value while preserving its code form."""
    try:
        return QuarterContext.parse(value).code
    except QuarterError as exc:
        raise argparse.ArgumentTypeError(str(exc)) from exc


def _parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="集团经济分析权属报表自动汇总")
    parser.add_argument(
        "--quarter",
        "-q",
        required=True,
        type=_quarter_argument,
        metavar="YYYYQn",
        help="当前季度，例如 2026Q2",
    )
    parser.add_argument(
        "--config",
        "-c",
        type=Path,
        default=None,
        help="配置文件路径；默认使用 main.py 同目录的 config.yaml",
    )
    parser.add_argument(
        "--template",
        "-t",
        dest="template",
        type=Path,
        default=None,
        help="指定汇总模板路径；未指定时使用 config.yaml 中的默认模板",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = _parse_args(argv)
    try:
        result = run(
            args.quarter,
            config_path=args.config,
            template_path=args.template,
        )
    except Exception:
        logging.getLogger(__name__).exception("汇总处理失败")
        return 1

    print(f"汇总文件: {result['output_path']}")
    print(f"处理日志: {result['log_path']}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
