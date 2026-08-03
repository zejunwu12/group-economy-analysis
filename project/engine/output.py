"""Summary workbook output helpers."""

import logging
import os
import re
import uuid
from datetime import datetime
from pathlib import Path

from openpyxl.workbook.workbook import Workbook


logger = logging.getLogger(__name__)

_WINDOWS_INVALID_FILENAME = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


class OutputSaveError(Exception):
    """The summary workbook could not be saved safely."""


def save_summary_workbook(
    workbook: Workbook,
    config: dict,
    *,
    artifact_label: str = "汇总文件",
) -> str:
    """Save a completed workbook and return its absolute path.

    The output directory is resolved relative to config.yaml.  A normal save
    replaces a same-name result atomically.  If Excel has locked that result,
    a timestamped copy is saved instead so the completed workbook is not lost.
    """
    output_dir = _resolve_output_dir(config)
    filename = _build_output_filename(config)
    output_dir.mkdir(parents=True, exist_ok=True)
    target = output_dir / filename

    try:
        _atomic_save(workbook, target)
    except PermissionError:
        fallback = _build_fallback_path(target)
        try:
            _atomic_save(workbook, fallback)
        except OSError as exc:
            raise OutputSaveError(
                f"{artifact_label}保存失败: {fallback}，原因: {exc}"
            ) from exc
        logger.warning(
            f"目标{artifact_label}正在使用，未覆盖: {target}；"
            f"已另存为: {fallback}"
        )
        return str(fallback)
    except OSError as exc:
        raise OutputSaveError(
            f"{artifact_label}保存失败: {target}，原因: {exc}"
        ) from exc

    logger.info(f"{artifact_label}已保存: {target}")
    return str(target)


def _resolve_output_dir(config: dict) -> Path:
    try:
        runtime = config["runtime"]
        configured_dir = runtime["output_dir"]
    except (KeyError, TypeError) as exc:
        raise OutputSaveError("缺少 runtime.output_dir 配置") from exc

    if not isinstance(configured_dir, str) or not configured_dir.strip():
        raise OutputSaveError("runtime.output_dir 必须是非空路径")

    config_dir = runtime.get("_config_dir", ".")
    return (Path(config_dir) / configured_dir).resolve()


def _build_output_filename(config: dict) -> str:
    try:
        runtime = config["runtime"]
        filename_template = runtime["output_filename"]
        quarter_label = config["quarter"]["label"]
    except (KeyError, TypeError) as exc:
        raise OutputSaveError("缺少季度标识或 runtime.output_filename 配置") from exc

    if not isinstance(filename_template, str) or not filename_template.strip():
        raise OutputSaveError("runtime.output_filename 必须是非空字符串")
    if not isinstance(quarter_label, str) or not quarter_label.strip():
        raise OutputSaveError("quarter.label 必须是非空字符串")

    try:
        filename = filename_template.format(quarter_label=quarter_label)
    except (KeyError, IndexError, ValueError) as exc:
        raise OutputSaveError(
            "runtime.output_filename 只能使用 {quarter_label} 占位符"
        ) from exc

    if Path(filename).name != filename:
        raise OutputSaveError("runtime.output_filename 不能包含目录")
    if _WINDOWS_INVALID_FILENAME.search(filename):
        raise OutputSaveError(f"输出文件名包含 Windows 不支持的字符: {filename}")
    if not filename.lower().endswith(".xlsx"):
        raise OutputSaveError("runtime.output_filename 必须以 .xlsx 结尾")
    return filename


def _atomic_save(workbook: Workbook, target: Path) -> None:
    """Save to a temporary xlsx and replace the target only after success."""
    temporary = target.with_name(
        f".{target.stem}.{uuid.uuid4().hex}.tmp{target.suffix}"
    )
    try:
        workbook.save(temporary)
        os.replace(temporary, target)
    except Exception:
        if temporary.exists():
            temporary.unlink()
        raise


def _build_fallback_path(target: Path) -> Path:
    """Create a collision-free timestamped filename next to the target."""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    candidate = target.with_name(f"{target.stem}_{timestamp}{target.suffix}")
    suffix = 1
    while candidate.exists():
        candidate = target.with_name(
            f"{target.stem}_{timestamp}_{suffix}{target.suffix}"
        )
        suffix += 1
    return candidate
