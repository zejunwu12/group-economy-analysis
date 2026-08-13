"""Logger module

配置方式：
    from logger import setup_logger
    setup_logger(log_level="INFO", output_dir="output/")
    # 之后各模块通过 logging.getLogger(__name__) 获取 logger 即可使用
"""

import logging
import os
from datetime import datetime
from pathlib import Path

from engine.comments import CommentCopyStats


SUPPORTED_LEVELS = {
    "DEBUG": logging.DEBUG,
    "INFO": logging.INFO,
    "WARNING": logging.WARNING,
    "ERROR": logging.ERROR,
    "CRITICAL": logging.CRITICAL,
}

_STEP_SEPARATOR = "=" * 18
_REPORT_SEPARATOR = "-" * 14
_CONSOLE_RECORD_ATTR = "console_summary"
_FILE_ONLY_RECORD_ATTR = "file_only"


class _ConciseConsoleFilter(logging.Filter):
    """Keep normal console output compact while retaining full file logs."""

    def __init__(self) -> None:
        super().__init__()
        self.suppressed_warning_count = 0

    def filter(self, record: logging.LogRecord) -> bool:
        if getattr(record, _FILE_ONLY_RECORD_ATTR, False):
            return False
        if getattr(record, _CONSOLE_RECORD_ATTR, False):
            return True
        if record.levelno >= logging.ERROR:
            return True
        if record.levelno >= logging.WARNING:
            self.suppressed_warning_count += 1
        return False


def setup_logger(
    log_level: str = "INFO",
    output_dir: str = "output/",
    file_log_level: str = "DEBUG",
    file_prefix: str = "summary",
) -> str:
    """初始化全局日志配置，同时输出简洁控制台日志和完整文件日志。

    Args:
        log_level: 控制台日志级别（DEBUG/INFO/WARNING/ERROR），默认 INFO
        output_dir: 日志文件输出目录
        file_log_level: 文件日志级别，默认 DEBUG，用于保留逐项追溯信息
        file_prefix: 日志文件名前缀，默认 ``summary``

    Returns:
        本次运行创建的日志文件绝对路径

    Raises:
        ValueError: 日志级别不受支持
    """
    normalized_level = log_level.strip().upper()
    if normalized_level not in SUPPORTED_LEVELS:
        supported = ", ".join(SUPPORTED_LEVELS)
        raise ValueError(
            f"不支持的日志级别 '{log_level}'，可选值: {supported}"
        )
    normalized_file_level = file_log_level.strip().upper()
    if normalized_file_level not in SUPPORTED_LEVELS:
        supported = ", ".join(SUPPORTED_LEVELS)
        raise ValueError(
            f"不支持的文件日志级别 '{file_log_level}'，可选值: {supported}"
        )
    if not file_prefix or any(
        character not in "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_-"
        for character in file_prefix
    ):
        raise ValueError("日志文件名前缀只能包含字母、数字、下划线和连字符")

    # 确保输出目录存在
    output_dir = os.path.abspath(output_dir)
    os.makedirs(output_dir, exist_ok=True)

    # 日志文件名：{file_prefix}_YYYYMMDD_HHMMSS.log
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = os.path.join(output_dir, f"{file_prefix}_{timestamp}.log")
    suffix = 1
    while os.path.exists(log_file):
        log_file = os.path.join(
            output_dir,
            f"{file_prefix}_{timestamp}_{suffix}.log",
        )
        suffix += 1

    # 日志格式
    fmt = "[%(asctime)s.%(msecs)03d] %(levelname)-8s %(name)s: %(message)s"
    datefmt = "%Y-%m-%d %H:%M:%S"

    # 获取 root logger 并关闭已有 handler（避免重复输出和文件句柄残留）
    root = logging.getLogger()
    for handler in root.handlers[:]:
        root.removeHandler(handler)
        handler.close()
    console_level = SUPPORTED_LEVELS[normalized_level]
    detailed_file_level = SUPPORTED_LEVELS[normalized_file_level]
    root.setLevel(min(console_level, detailed_file_level))

    # 控制台 handler
    console = logging.StreamHandler()
    console.set_name("summary-console")
    console.setLevel(console_level)
    console.setFormatter(logging.Formatter(fmt, datefmt))
    if console_level > logging.DEBUG:
        console.addFilter(_ConciseConsoleFilter())
    root.addHandler(console)

    # 文件 handler
    file_handler = logging.FileHandler(log_file, encoding="utf-8")
    file_handler.set_name("summary-file")
    file_handler.setLevel(detailed_file_level)
    file_handler.setFormatter(logging.Formatter(fmt, datefmt))
    root.addHandler(file_handler)

    # 记录启动信息
    logging.getLogger(__name__).debug(
        "日志系统初始化完成: 控制台=%s，文件=%s，路径=%s",
        normalized_level,
        normalized_file_level,
        log_file,
    )
    return log_file


def log_workflow_step(
    step_number: int,
    total_steps: int,
    title: str,
    *,
    step_logger: logging.Logger | None = None,
) -> None:
    """输出一个醒目的主流程步骤标题。"""
    active_logger = step_logger or logging.getLogger("processing.workflow")
    active_logger.info(
        "%s [步骤 %s/%s] %s %s",
        _STEP_SEPARATOR,
        step_number,
        total_steps,
        title,
        _STEP_SEPARATOR,
        extra={_CONSOLE_RECORD_ATTR: True},
    )


def log_report_step(
    position: int,
    total_reports: int,
    report_id: int,
    sheet_name: str,
    *,
    step_logger: logging.Logger | None = None,
) -> None:
    """输出报表写入阶段的子步骤标题。"""
    active_logger = step_logger or logging.getLogger("processing.workflow")
    display_name = _display_sheet_name(report_id, sheet_name)
    active_logger.info(
        "%s [报表 %s/%s] 报表%s %s %s",
        _REPORT_SEPARATOR,
        position,
        total_reports,
        report_id,
        display_name,
        _REPORT_SEPARATOR,
        extra={_FILE_ONLY_RECORD_ATTR: True},
    )


def get_suppressed_console_warning_count() -> int:
    """Return warnings hidden from the concise console during this run."""
    for handler in logging.getLogger().handlers:
        if handler.get_name() != "summary-console":
            continue
        for log_filter in handler.filters:
            if isinstance(log_filter, _ConciseConsoleFilter):
                return log_filter.suppressed_warning_count
    return 0


def log_ownership_check_details(
    config: dict,
    ownership_data: dict,
    header_mismatches: dict[tuple[str, int], str] | None = None,
    *,
    detail_logger: logging.Logger | None = None,
    console_summary: bool = False,
) -> None:
    """输出权属文件加载及格式检查详情。"""
    active_logger = detail_logger or logging.getLogger("processing.ownership")
    header_mismatches = header_mismatches or {}
    ownership_statuses = _build_ownership_statuses(config, ownership_data)
    ownership_counts = _count_statuses(ownership_statuses.values())
    summary_extra = {
        _CONSOLE_RECORD_ATTR if console_summary else _FILE_ONLY_RECORD_ATTR: True
    }

    active_logger.info(
        "[权属文件加载] 已加载 %s 个，未加载 %s 个，无需加载 %s 个",
        ownership_counts["success"],
        ownership_counts["failed"],
        ownership_counts["skipped"],
        extra=summary_extra,
    )
    for owner_key, status in ownership_statuses.items():
        if status["status"] == "success":
            continue
        message = f"  {owner_key}: {status['label']}"
        if status["detail"]:
            message += f"（{status['detail']}）"
        _log_by_status(
            active_logger,
            status["status"],
            message,
            file_only=True,
        )

    report_count = len(config["runtime"].get("reports_to_run", range(1, 9)))
    header_check_total = len(ownership_data) * report_count
    header_check_passed = max(header_check_total - len(header_mismatches), 0)
    if header_mismatches:
        active_logger.warning(
            f"  [权属表格式检查] 通过 {header_check_passed} 项，"
            f"需核对 {len(header_mismatches)} 项",
            extra=summary_extra,
        )
        active_logger.warning(
            "    需核对问题明细：",
            extra={_FILE_ONLY_RECORD_ATTR: True},
        )
        for problem_number, ((owner_key, report_id), reason) in enumerate(
            sorted(header_mismatches.items()),
            start=1,
        ):
            for message in _format_header_mismatch_warning(
                config,
                owner_key,
                report_id,
                reason,
                problem_number,
            ):
                active_logger.warning(
                    message,
                    extra={_FILE_ONLY_RECORD_ATTR: True},
                )
    else:
        active_logger.info(
            f"  [权属表格式检查] 通过 {header_check_passed} 项，需核对 0 项",
            extra=summary_extra,
        )


def log_processing_summary(
    config: dict,
    ownership_data: dict,
    report_results: dict,
    total_validation: dict | None = None,
    anomaly_report: dict | None = None,
    output_path: str | None = None,
    report_errors: dict | None = None,
    comment_stats: CommentCopyStats | None = None,
    summary_logger: logging.Logger | None = None,
    header_mismatches: dict[tuple[str, int], str] | None = None,
) -> dict:
    """输出一次汇总处理的完整摘要，并返回摘要统计。

    Args:
        config: ConfigLoader.load() 返回的配置。
        ownership_data: 已成功加载的权属数据，结构与 reader.load_ownership_files()
            的返回值一致。
        report_results: 各报表处理函数的返回结果，键可为 1 或 ``report1``。
        header_mismatches: 权属表格式异常，键为 ``(权属名称, 报表编号)``，
            值为异常原因。
        total_validation: validator.validate_totals() 返回的报告。
        anomaly_report: validator.detect_anomalies() 返回的报告。
        output_path: save_summary_workbook() 返回的实际保存路径。
        report_errors: 可选的报表错误，键可为 1 或 ``report1``。
        summary_logger: 可选日志记录器，默认使用 ``processing.summary``。
    """
    active_logger = summary_logger or logging.getLogger("processing.summary")
    report_errors = report_errors or {}
    header_mismatches = header_mismatches or {}
    ownership_statuses = _build_ownership_statuses(config, ownership_data)
    report_statuses = _build_report_statuses(
        config,
        report_results,
        report_errors,
    )

    separator = "=" * 24
    active_logger.info(
        f"{separator} 运行结果摘要 {separator}",
        extra={_CONSOLE_RECORD_ATTR: True},
    )
    active_logger.info(
        f"汇总周期: {config['quarter']['label']}",
        extra={_FILE_ONLY_RECORD_ATTR: True},
    )

    log_ownership_check_details(
        config,
        ownership_data,
        header_mismatches,
        detail_logger=active_logger,
        console_summary=True,
    )

    report_counts = _count_statuses(report_statuses.values())
    active_logger.info(
        "[报表处理结果] 完成 %s 张，失败 %s 张，未执行 %s 张",
        report_counts["success"],
        report_counts["failed"],
        report_counts["skipped"],
        extra={_CONSOLE_RECORD_ATTR: True},
    )
    for report_id, status in report_statuses.items():
        sheet_name = _display_sheet_name(report_id, status["sheet_name"])
        message = (
            f"  报表{report_id} {sheet_name}: "
            f"{status['label']}，{status['detail']}"
        )
        _log_by_status(
            active_logger,
            status["status"],
            message,
            file_only=True,
        )

    active_logger.info(
        "[合计值校验]",
        extra={_CONSOLE_RECORD_ATTR: True},
    )
    if total_validation is None:
        active_logger.info(
            "  未执行",
            extra={_CONSOLE_RECORD_ATTR: True},
        )
    else:
        active_logger.info(
            f"  通过 {total_validation.get('passed', 0)} 项，"
            f"需核对 {total_validation.get('warnings', 0)} 项，"
            f"未校验 {total_validation.get('skipped', 0)} 项",
            extra={_CONSOLE_RECORD_ATTR: True},
        )
        warning_details = [
            detail
            for detail in total_validation.get("details", [])
            if detail.get("status") == "warning"
        ]
        if warning_details:
            active_logger.warning(
                "  需核对问题明细：",
                extra={_FILE_ONLY_RECORD_ATTR: True},
            )
            for problem_number, detail in enumerate(warning_details, start=1):
                for message in _format_total_warning(detail, problem_number):
                    active_logger.warning(
                        message,
                        extra={_FILE_ONLY_RECORD_ATTR: True},
                    )

    active_logger.info(
        "[数据完整性和异常值检测]",
        extra={_CONSOLE_RECORD_ATTR: True},
    )
    if anomaly_report is None:
        active_logger.info(
            "  未执行",
            extra={_CONSOLE_RECORD_ATTR: True},
        )
    else:
        active_logger.info(
            f"  共发现 {anomaly_report.get('total', 0)} 项："
            f"负面积 {anomaly_report.get('negative_area', 0)} 项，"
            f"超过面积阈值 {anomaly_report.get('oversized_area', 0)} 项，"
            f"关键字段均为空 {anomaly_report.get('empty_key_data', 0)} 项",
            extra={_CONSOLE_RECORD_ATTR: True},
        )
        anomaly_details = anomaly_report.get("details", [])
        if anomaly_details:
            active_logger.warning(
                "  异常明细：",
                extra={_FILE_ONLY_RECORD_ATTR: True},
            )
            for problem_number, detail in enumerate(anomaly_details, start=1):
                active_logger.warning(
                    _format_anomaly_warning(detail, problem_number),
                    extra={_FILE_ONLY_RECORD_ATTR: True},
                )

    active_logger.info(
        "[批注处理]",
        extra={_CONSOLE_RECORD_ATTR: True},
    )
    active_logger.info(
        f"  模板原有批注已清除 "
        f"{0 if comment_stats is None else comment_stats.template_cleared} 条；"
        f"权属批注已复制 {0 if comment_stats is None else comment_stats.copied} 条",
        extra={_CONSOLE_RECORD_ATTR: True},
    )
    comment_details = [] if comment_stats is None else comment_stats.details
    if comment_details:
        active_logger.info(
            "  权属批注复制明细：",
            extra={_FILE_ONLY_RECORD_ATTR: True},
        )
        for comment_number, detail in enumerate(comment_details, start=1):
            for message in _format_comment_detail(detail, comment_number):
                active_logger.info(
                    message,
                    extra={_FILE_ONLY_RECORD_ATTR: True},
                )

    active_logger.info(
        "[生成文件]",
        extra={_FILE_ONLY_RECORD_ATTR: True},
    )
    if output_path:
        active_logger.info(
            f"  {Path(output_path).resolve()}",
            extra={_FILE_ONLY_RECORD_ATTR: True},
        )
    else:
        active_logger.warning(
            "  未生成汇总文件",
            extra={_CONSOLE_RECORD_ATTR: True},
        )
    active_logger.info(
        f"{separator} 摘要结束 {separator}",
        extra={_FILE_ONLY_RECORD_ATTR: True},
    )

    summary = {
        "ownership": _count_statuses(ownership_statuses.values()),
        "header_mismatches": len(header_mismatches),
        "reports": _count_statuses(report_statuses.values()),
        "total_validation": {
            "passed": 0 if total_validation is None else total_validation.get("passed", 0),
            "warnings": 0 if total_validation is None else total_validation.get("warnings", 0),
            "skipped": 0 if total_validation is None else total_validation.get("skipped", 0),
        },
        "anomalies": 0 if anomaly_report is None else anomaly_report.get("total", 0),
        "template_comments_cleared": (
            0 if comment_stats is None else comment_stats.template_cleared
        ),
        "comments_copied": 0 if comment_stats is None else comment_stats.copied,
        "output_path": output_path,
    }
    _flush_handlers(active_logger)
    return summary


def _build_ownership_statuses(config: dict, ownership_data: dict) -> dict:
    statuses = {}
    for owner_key, entry in config["ownership_files"].items():
        filename = entry.get("file")
        if filename is None:
            statuses[owner_key] = {
                "status": "skipped",
                "label": "无需加载",
                "detail": "配置中未指定独立数据文件",
            }
        elif owner_key in ownership_data:
            loaded_name = ownership_data[owner_key].get("filename", filename)
            statuses[owner_key] = {
                "status": "success",
                "label": "已加载",
                "detail": loaded_name,
            }
        else:
            statuses[owner_key] = {
                "status": "failed",
                "label": "未加载",
                "detail": f"未找到文件：{filename}",
            }
    return statuses


def _build_report_statuses(
    config: dict,
    report_results: dict,
    report_errors: dict,
) -> dict:
    statuses = {}
    reports_to_run = config["runtime"].get("reports_to_run", list(range(1, 9)))
    for report_id in reports_to_run:
        report_key = f"report{report_id}"
        sheet_name = config["reports"][report_key]["sheet_name"]
        error = _get_by_report_key(report_errors, report_id)
        result = _get_by_report_key(report_results, report_id)
        if error is not None:
            statuses[report_id] = {
                "status": "failed",
                "label": "处理失败",
                "detail": str(error),
                "sheet_name": sheet_name,
            }
        elif result is not None:
            statuses[report_id] = {
                "status": "success",
                "label": "处理完成",
                "detail": _format_report_result(report_id, result),
                "sheet_name": sheet_name,
            }
        else:
            statuses[report_id] = {
                "status": "skipped",
                "label": "未执行",
                "detail": "没有处理结果",
                "sheet_name": sheet_name,
            }
    return statuses


def _get_by_report_key(values: dict, report_id: int):
    if report_id in values:
        return values[report_id]
    return values.get(f"report{report_id}")


def _format_report_result(report_id: int, result: object) -> str:
    if isinstance(result, bool):
        return "处理完成"
    if isinstance(result, int):
        return f"写入 {result} 行"
    if not isinstance(result, dict):
        return "处理完成"
    if report_id == 5:
        return (
            f"左侧 {result.get('left_count', 0)} 条，"
            f"右侧 {result.get('right_count', 0)} 条，"
            f"数据区至第 {result.get('data_end_row', '?')} 行"
        )
    if report_id == 7:
        written_cells = sum(
            item.get("written_cells", 0)
            for item in result.values()
            if isinstance(item, dict)
        )
        return f"写入 {len(result)} 个园区，共 {written_cells} 个单元格"
    if report_id == 8:
        return (
            f"写入 {result.get('record_count', 0)} 条明细，"
            f"数据区至第 {result.get('data_end_row', '?')} 行"
        )
    if "written_count" in result:
        return f"写入 {result['written_count']} 行"
    return "处理完成"


def _display_sheet_name(report_id: int, sheet_name: str) -> str:
    """去掉工作表名称中重复的报表编号，仅优化日志显示。"""
    prefix = f"报表{report_id} "
    if sheet_name.startswith(prefix):
        return sheet_name[len(prefix):]
    return sheet_name


def _format_total_warning(detail: dict, problem_number: int) -> list[str]:
    """将一项合计值问题拆成便于人工核对的多行说明。"""
    report_id = detail.get("report_id")
    sheet_name = _display_sheet_name(report_id, detail.get("sheet_name", ""))
    messages = [
        f"    [问题{problem_number}] 报表{report_id} {sheet_name}｜"
        f"{detail.get('column')}列",
        f"      原因：{detail.get('reason', '合计值校验异常')}",
    ]

    formula = detail.get("formula")
    formula_range = detail.get("formula_range")
    if formula is not None or formula_range is not None:
        formula_message = f"      合计公式：{formula or '未设置'}"
        if formula_range is not None:
            formula_message += f"｜引用范围：{formula_range}"
        messages.append(formula_message)

    value_parts = []
    if "formula_sum" in detail:
        value_parts.append(
            "公式范围合计：" + _format_log_value(detail["formula_sum"])
        )
    if "expected_sum" in detail:
        value_parts.append(
            "完整数据区合计：" + _format_log_value(detail["expected_sum"])
        )
    if "difference" in detail:
        value_parts.append(
            "差额（公式范围合计-完整数据区合计）："
            + _format_log_value(detail["difference"])
        )
    if value_parts:
        messages.append("      " + "｜".join(value_parts))

    return messages


def _format_header_mismatch_warning(
    config: dict,
    owner_key: str,
    report_id: int,
    reason: str,
    problem_number: int,
) -> list[str]:
    """将一项权属表格式异常拆成便于人工核对的多行说明。"""
    sheet_name = config["reports"][f"report{report_id}"]["sheet_name"]
    display_name = _display_sheet_name(report_id, sheet_name)
    return [
        f"      [问题{problem_number}] 报表{report_id} {display_name}｜"
        f"权属：{owner_key}",
        f"        原因：{reason}",
        "        处理结果：该权属本报表未写入，请人工复核",
    ]


def _format_anomaly_warning(detail: dict, problem_number: int) -> str:
    """生成一项数据完整性或异常值问题的摘要说明。"""
    report_id = detail.get("report_id")
    sheet_name = _display_sheet_name(report_id, detail.get("sheet_name", ""))
    return (
        f"    [问题{problem_number}] 报表{report_id} {sheet_name}｜"
        f"位置：{detail.get('cell')}｜"
        f"{detail.get('message', '数据异常')}｜"
        f"标识值：{_format_log_value(detail.get('value'))}"
    )


def _format_log_value(value: object) -> str:
    """格式化摘要中的数值，保留精度并增加千位分隔符。"""
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return str(value)
    if isinstance(value, int):
        return f"{value:,}"
    if value == 0:
        return "0"
    return f"{value:,.10f}".rstrip("0").rstrip(".")


def _format_comment_detail(detail, comment_number: int) -> list[str]:
    """将一条权属批注复制记录格式化为摘要中的多行说明。"""
    sheet_name = _display_sheet_name(detail.report_id, detail.target_sheet)
    comment_text = " / ".join(str(detail.text).splitlines()) or "（空白）"
    return [
        f"    [批注{comment_number}] 报表{detail.report_id} {sheet_name}｜"
        f"权属：{detail.owner_key}",
        f"      来源：{detail.source_sheet}!{detail.source_cell}｜"
        f"目标：{detail.target_sheet}!{detail.target_cell}｜作者：{detail.author}",
        f"      内容：{comment_text}",
    ]


def _log_by_status(
    active_logger: logging.Logger,
    status: str,
    message: str,
    *,
    file_only: bool = False,
) -> None:
    extra = {_FILE_ONLY_RECORD_ATTR: True} if file_only else None
    if status == "failed":
        active_logger.error(message, extra=extra)
    elif status == "skipped":
        active_logger.warning(message, extra=extra)
    else:
        active_logger.info(message, extra=extra)


def _count_statuses(statuses) -> dict:
    result = {"success": 0, "failed": 0, "skipped": 0}
    for status in statuses:
        result[status["status"]] += 1
    return result


def _flush_handlers(active_logger: logging.Logger) -> None:
    handlers = set(active_logger.handlers)
    current = active_logger
    while current.propagate and current.parent is not None:
        current = current.parent
        handlers.update(current.handlers)
    for handler in handlers:
        handler.flush()
