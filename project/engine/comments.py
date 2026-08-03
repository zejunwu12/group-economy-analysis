"""权属来源批注向汇总表复制的公共逻辑。"""

import logging
from copy import copy
from dataclasses import dataclass, field

from openpyxl.cell.cell import Cell
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet


@dataclass
class CommentCopyDetail:
    """一条最终保留在汇总工作簿中的权属批注复制记录。"""

    report_id: int
    owner_key: str
    source_sheet: str
    source_cell: str
    target_sheet: str
    target_cell: str
    author: str
    text: str


@dataclass
class CommentCopyStats:
    """记录模板批注清理数量及权属批注复制明细。"""

    copied: int = 0
    template_cleared: int = 0
    details: list[CommentCopyDetail] = field(default_factory=list)

    def record(self, detail: CommentCopyDetail) -> None:
        """登记一条成功复制的权属批注。"""
        self.copied += 1
        self.details.append(detail)

    def checkpoint(self) -> int:
        """返回当前批注明细位置，供报表失败时回滚。"""
        return len(self.details)

    def rollback(self, checkpoint: int) -> int:
        """回滚检查点之后的批注明细并返回撤销数量。"""
        if checkpoint < 0 or checkpoint > len(self.details):
            raise ValueError(f"无效的批注统计检查点: {checkpoint}")
        removed = len(self.details) - checkpoint
        if removed:
            del self.details[checkpoint:]
            self.copied -= removed
        return removed


def clear_template_comments(
    workbook: Workbook,
    *,
    stats: CommentCopyStats | None = None,
) -> int:
    """清除模板中已有批注，避免遗留说明混入本次汇总结果。"""
    cleared = 0
    logger = logging.getLogger(__name__)
    for worksheet in workbook.worksheets:
        for row in worksheet.iter_rows():
            for cell in row:
                if cell.comment is None:
                    continue
                author = cell.comment.author or "未署名"
                logger.debug(
                    "  [模板批注已清除] %s!%s；原作者=%s",
                    worksheet.title,
                    cell.coordinate,
                    author,
                )
                cell.comment = None
                cleared += 1

    if stats is not None:
        stats.template_cleared += cleared
    logger.info("模板原有批注清理完成: 共清除 %s 条", cleared)
    return cleared


def copy_source_comment(
    source_cell: Cell,
    target_cell: Cell,
    *,
    report_id: int,
    owner_key: str,
    source_ws: Worksheet,
    target_ws: Worksheet,
    stats: CommentCopyStats | None = None,
) -> bool:
    """将来源批注完整复制到目标单元格，并记录可追溯日志。

    调用方只应在来源单元格确实参与数据写入、且目标单元格可写时调用。
    使用浅拷贝避免来源与目标共享同一个批注对象，作者和正文均按来源
    原样保留，原始权属文件不作修改。
    """
    source_comment = source_cell.comment
    if source_comment is None:
        return False

    author = source_comment.author or "未署名"
    text = source_comment.text or ""
    target_cell.comment = copy(source_comment)
    if stats is not None:
        stats.record(
            CommentCopyDetail(
                report_id=report_id,
                owner_key=owner_key,
                source_sheet=source_ws.title,
                source_cell=source_cell.coordinate,
                target_sheet=target_ws.title,
                target_cell=target_cell.coordinate,
                author=author,
                text=text,
            )
        )
    logging.getLogger(__name__).debug(
        "  [批注已复制] 报表%s 权属 '%s' %s!%s → %s!%s；"
        "作者=%s；内容=%s",
        report_id,
        owner_key,
        source_ws.title,
        source_cell.coordinate,
        target_ws.title,
        target_cell.coordinate,
        author,
        text,
    )
    return True
