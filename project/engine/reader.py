"""Template and data reader"""

import logging
import os
from pathlib import Path
import openpyxl
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.workbook.workbook import Workbook

logger = logging.getLogger(__name__)


class TemplateMismatchError(Exception):
    """模板与配置不一致异常"""
    pass


def load_template(
    config: dict,
    template_path: str | os.PathLike[str] | None = None,
) -> Workbook:
    """加载汇总表模板。

    Args:
        config: 由 ConfigLoader.load() 返回的配置字典
        template_path: 可选的运行时模板路径；未提供时使用
            ``runtime.template_path``，相对于 config.yaml 所在目录解析。

    Returns:
        openpyxl Workbook 对象

    Raises:
        FileNotFoundError: 模板文件不存在
    """
    if template_path is None:
        configured_path = config["runtime"]["template_path"]
        config_dir = config["runtime"].get("_config_dir", "")
        # 配置中的模板路径相对于 config.yaml 所在目录。
        full_path = Path(config_dir, configured_path).resolve()
    else:
        if not str(template_path).strip():
            raise ValueError("运行时模板路径不能为空")
        # 命令行指定的路径相对于当前命令执行目录。
        full_path = Path(template_path).expanduser().resolve()

    if not full_path.is_file():
        raise FileNotFoundError(f"模板文件不存在: {full_path}")

    resolved_path = str(full_path)
    logger.info(f"加载模板: {resolved_path}")
    wb = openpyxl.load_workbook(
        resolved_path,
        read_only=False,
        data_only=False,
        keep_links=True,
    )

    # 输出工作表信息
    logger.debug(f"模板包含 {len(wb.sheetnames)} 个工作表: {wb.sheetnames}")
    for name in wb.sheetnames:
        ws = wb[name]
        merged_count = len(ws.merged_cells.ranges)
        logger.debug(
            f"  {name}: {ws.max_row}行 × {ws.max_column}列, "
            f"合并单元格 {merged_count} 处"
        )

    return wb


def load_data_workbook(file_path: str) -> Workbook:
    """加载权属数据文件。

    Args:
        file_path: 权属文件的完整路径

    Returns:
        openpyxl Workbook 对象

    Raises:
        FileNotFoundError: 文件不存在
    """
    file_path = os.path.abspath(file_path)
    if not os.path.isfile(file_path):
        raise FileNotFoundError(f"权属文件不存在: {file_path}")

    wb = openpyxl.load_workbook(
        file_path,
        read_only=False,
        data_only=True,
        keep_links=True,
    )
    return wb


def validate_template(
    wb: Workbook,
    config: dict,
    report_ids: tuple[int, ...] | list[int] | None = None,
) -> list[str]:
    """校验模板中 row_mapping 声明的行号与实际内容是否一致。

    对固定行数报表（1/2/3/4/6），检查 B 列值是否与配置的 row_mapping 匹配。
    对报表7，检查每个子表 A 列（园区名称）是否与配置匹配。
    报表5/8 无 row_mapping，跳过。

    模板中的名称可能包含空格、换行或制表符，因此比对前会移除空白字符；
    归一化后仍采用精确匹配，名称差异应在 row_mapping 中明确配置。

    Args:
        wb: 已加载的模板 Workbook
        config: 配置字典
        report_ids: 仅校验指定报表；None 时保持兼容，校验全部适用报表

    Returns:
        校验通过的报表名称列表

    Raises:
        TemplateMismatchError: 发现不一致时立即抛出
    """
    reports = config["reports"]
    passed = []
    selected_report_ids = (
        set(report_ids) if report_ids is not None else {1, 2, 3, 4, 5, 6, 7, 8}
    )

    for report_id in (1, 2, 3, 4, 6):
        if report_id not in selected_report_ids:
            continue
        report_key = f"report{report_id}"
        rc = reports[report_key]
        sheet_name = rc["sheet_name"]
        if sheet_name not in wb.sheetnames:
            raise TemplateMismatchError(f"模板中不存在工作表: '{sheet_name}'")
        ws = wb[sheet_name]
        b_col = rc["b_col"]

        for row_num, expected_name in rc["row_mapping"].items():
            cell_ref = f"{b_col}{row_num}"
            actual_value = ws[cell_ref].value

            # 处理合并单元格中 None 值的情况
            if actual_value is None:
                actual_value = _get_merged_cell_value(ws, cell_ref)

            if actual_value is None:
                raise TemplateMismatchError(
                    f"{sheet_name} 第{row_num}行 {b_col}列为空，"
                    f"期望值: '{expected_name}'"
                )

            if not _names_match(expected_name, str(actual_value)):
                raise TemplateMismatchError(
                    f"{sheet_name} 第{row_num}行不匹配: "
                    f"期望 '{expected_name}'，实际 '{str(actual_value).strip()}'"
                )

        logger.debug(f"  {sheet_name}: row_mapping 校验通过 ({len(rc['row_mapping'])}行)")
        passed.append(sheet_name)

    # 校验报表7（园区）
    if 7 not in selected_report_ids:
        logger.info(f"模板一致性校验通过: 共 {len(passed)} 个数据区")
        return passed

    rc7 = reports["report7"]
    sheet_name = rc7["sheet_name"]
    if sheet_name not in wb.sheetnames:
        raise TemplateMismatchError(f"模板中不存在工作表: '{sheet_name}'")
    ws = wb[sheet_name]

    for sub_table in rc7["sub_tables"]:
        sub_name = sub_table["name"]
        for row_num, expected_name in sub_table["row_mapping"].items():
            cell_ref = f"A{row_num}"
            actual_value = ws[cell_ref].value

            if actual_value is None:
                actual_value = _get_merged_cell_value(ws, cell_ref)

            if actual_value is None:
                raise TemplateMismatchError(
                    f"{sheet_name} [{sub_name}] 第{row_num}行 A列为空，"
                    f"期望值: '{expected_name}'"
                )

            if not _names_match(expected_name, str(actual_value)):
                raise TemplateMismatchError(
                    f"{sheet_name} [{sub_name}] 第{row_num}行不匹配: "
                    f"期望 '{expected_name}'，实际 '{str(actual_value).strip()}'"
                )

        logger.debug(f"  {sheet_name} [{sub_name}]: row_mapping 校验通过 "
                     f"({len(sub_table['row_mapping'])}个园区)")
        passed.append(f"{sheet_name} [{sub_name}]")

    logger.info(f"模板一致性校验通过: 共 {len(passed)} 个数据区")
    return passed


def _normalize_name(name: str) -> str:
    """归一化名称：移除所有空白字符（空格、换行、制表符等）。"""
    return "".join(name.split())


def _names_match(expected: str, actual: str) -> bool:
    """判断两个名称是否匹配（归一化后精确匹配）。

    去除所有空白字符（空格、换行、制表符等）后比对。
    模板中名称有差异的，应在各报表的 row_mapping 中分别配置实际名称。
    """
    return _normalize_name(expected) == _normalize_name(actual)


def _get_merged_cell_value(ws: Worksheet, cell_ref: str) -> str | None:
    """如果单元格属于合并区域且为空，尝试从合并区域左上角获取值。

    Args:
        ws: 工作表
        cell_ref: 单元格引用，如 "B7"

    Returns:
        合并区域左上角的值，如果不在合并区域中则返回 None
    """
    cell = ws[cell_ref]
    for merged_range in ws.merged_cells.ranges:
        if cell.coordinate in merged_range:
            # 获取合并区域左上角单元格的值
            top_left = merged_range.min_row, merged_range.min_col
            top_left_cell = ws.cell(row=top_left[0], column=top_left[1])
            return top_left_cell.value
    return None


def load_ownership_files(config: dict) -> dict:
    """扫描数据目录，按精确文件名加载所有权属文件。

    Args:
        config: 配置字典

    Returns:
        {ownership_key: {"workbook": wb, "file": path, "filename": name}}
        不包含 file 为 null 的条目

    Raises:
        FileNotFoundError: 数据目录不存在，或没有任何可加载的权属文件
    """
    config_dir = config["runtime"].get("_config_dir", "")
    data_dir = os.path.abspath(
        os.path.normpath(
            os.path.join(config_dir, config["runtime"]["data_dir"])
        )
    )

    if not os.path.isdir(data_dir):
        raise FileNotFoundError(f"数据目录不存在: {data_dir}")

    logger.info(f"数据目录: {data_dir}")

    ownership_files = config["ownership_files"]
    expected_filenames = {
        entry["file"]
        for entry in ownership_files.values()
        if entry["file"] is not None
    }
    available_filenames = {
        entry.name
        for entry in os.scandir(data_dir)
        if entry.is_file()
        and entry.name.lower().endswith(".xlsx")
        and not entry.name.startswith("~$")
    }

    for filename in sorted(available_filenames - expected_filenames):
        logger.warning(f"未识别的权属文件，已跳过: {filename}")

    result = {}
    warnings = []
    load_errors = []

    for owner_key, entry in ownership_files.items():
        filename = entry["file"]

        # 无独立数据文件的跳过
        if filename is None:
            logger.debug(f"  跳过 {owner_key}: 无独立数据文件")
            continue

        file_path = os.path.join(data_dir, filename)

        if filename not in available_filenames:
            msg = (
                "未找到权属文件（文件不存在或文件名与配置不一致）："
                f"{file_path}"
            )
            warnings.append(msg)
            continue

        logger.debug(f"  加载 {owner_key}: {filename}")

        try:
            wb = load_data_workbook(file_path)
        except Exception as e:
            msg = f"加载失败 {owner_key} ({filename}): {e}"
            load_errors.append(msg)
            logger.error(msg)
            continue

        result[owner_key] = {
            "workbook": wb,
            "file": file_path,
            "filename": filename,
        }

    for msg in warnings:
        logger.warning(msg)

    if len(result) == 0:
        problem_count = len(warnings) + len(load_errors)
        raise FileNotFoundError(
            f"数据目录中无任何可加载文件，共 {problem_count} 个问题"
        )

    logger.info(f"已加载权属文件: {len(result)} 个")
    return result


def close_ownership_files(ownership_data: dict) -> None:
    """关闭所有权属文件的 workbook。

    Args:
        ownership_data: load_ownership_files() 的返回值
    """
    for owner_key, data in ownership_data.items():
        try:
            data["workbook"].close()
            logger.debug(f"  关闭 {owner_key}")
        except Exception as e:
            logger.warning(f"关闭 {owner_key} 失败: {e}")
