"""Excel 名单读取与中奖结果导出。"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Sequence

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

from .models import PrizeTier, Winner

# 用于猜测“姓名列/标题行”的常见表头关键词
NAME_HEADER_HINTS = (
    "姓名",
    "名字",
    "名单",
    "员工",
    "人员",
    "参与人",
    "昵称",
    "name",
    "username",
)

HEADER_SCAN_ROWS = 10


@dataclass
class ColumnOption:
    index: int  # 工作表中的 0 基列号
    letter: str
    header: str | None
    values: list[str]

    @property
    def label(self) -> str:
        text = f"{self.letter} 列"
        if self.header:
            text += f" · {self.header}"
        return f"{text}（{len(self.values)} 条）"


@dataclass
class SheetPreview:
    sheet: str
    columns: list[ColumnOption]
    has_header: bool


def cell_text(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def list_sheets(path: str | Path) -> list[str]:
    """返回工作簿中的所有工作表名。"""
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        return list(wb.sheetnames)
    finally:
        wb.close()


def inspect_sheet(path: str | Path, sheet: str) -> SheetPreview:
    """读取工作表并识别标题行、姓名列。

    在开头若干行中寻找含“姓名/name”等关键词的标题行：
    - 找到：该行作为表头，其后为数据；
    - 找不到：不跳过任何行，按纯数据表处理。
    """
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        if sheet not in wb.sheetnames:
            raise ValueError(f"工作表「{sheet}」不存在")
        ws = wb[sheet]
        rows = [[cell_text(v) for v in row] for row in ws.iter_rows(values_only=True)]
    finally:
        wb.close()

    while rows and not any(rows[-1]):
        rows.pop()

    header_row: list[str] | None = None
    header_index = -1
    for idx, row in enumerate(rows[:HEADER_SCAN_ROWS]):
        if any(hint in cell.lower() for cell in row for hint in NAME_HEADER_HINTS):
            header_row = row
            header_index = idx
            break

    if header_row is not None:
        data_rows = rows[header_index + 1 :]
    else:
        data_rows = [row for row in rows if any(row)]

    width = max((len(row) for row in data_rows), default=0)
    columns: list[ColumnOption] = []
    for col in range(width):
        values = [row[col] for row in data_rows if col < len(row) and row[col]]
        if not values:
            continue
        header = None
        if header_row is not None and col < len(header_row):
            header = header_row[col] or None
        columns.append(
            ColumnOption(
                index=col,
                letter=get_column_letter(col + 1),
                header=header,
                values=values,
            )
        )

    return SheetPreview(sheet=sheet, columns=columns, has_header=header_row is not None)


def default_column(preview: SheetPreview) -> int | None:
    """优先选表头命中关键词的列，否则选第一条有数据的列。"""
    for pos, col in enumerate(preview.columns):
        if col.header and any(h in col.header.lower() for h in NAME_HEADER_HINTS):
            return pos
    return 0 if preview.columns else None


def read_names(path: str | Path, sheet: str, column_index: int) -> list[str]:
    """按 `inspect_sheet` 的列序号读取姓名，保留重复姓名（可能是同名不同人）。"""
    preview = inspect_sheet(path, sheet)
    if not 0 <= column_index < len(preview.columns):
        raise ValueError("选择的列不存在")
    return list(preview.columns[column_index].values)


def export_results(
    path: str | Path,
    winners: Sequence[Winner],
    tiers: Sequence[PrizeTier],
) -> Path:
    """导出中奖名单与各奖项统计到新的 xlsx 文件。"""
    path = Path(path)
    wb = Workbook()

    ws = wb.active
    ws.title = "中奖名单"
    ws.append(["编号", "奖项", "奖品", "姓名", "中奖时间"])
    for winner in winners:
        ws.append(
            [
                winner.number,
                winner.tier_name,
                winner.prize,
                winner.name,
                winner.drawn_at.strftime("%Y-%m-%d %H:%M:%S"),
            ]
        )
    for col, width in zip("ABCDE", (10, 16, 22, 18, 22)):
        ws.column_dimensions[col].width = width

    summary = wb.create_sheet("奖品统计")
    summary.append(["奖项", "奖品", "计划人数", "已抽人数", "状态"])
    by_tier: dict[str, int] = {}
    for winner in winners:
        by_tier[winner.tier_id] = by_tier.get(winner.tier_id, 0) + 1
    for tier in tiers:
        drawn = by_tier.get(tier.id, 0)
        summary.append(
            [
                tier.name,
                tier.prize,
                tier.count,
                drawn,
                "已完成" if drawn >= tier.count else "未完成",
            ]
        )
    summary.append([])
    summary.append(["导出时间", datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
    for col, width in zip("ABCDE", (16, 22, 12, 12, 12)):
        summary.column_dimensions[col].width = width

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path
