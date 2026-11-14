"""从 Excel 导入奖项设置。"""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from .excel_io import cell_text
from .models import PrizeTier

TIER_NAME_HINTS = ("奖项名称", "奖项", "等级", "奖级", "level")
PRIZE_HINTS = ("奖品", "礼品", "物品", "奖金", "prize")
COUNT_HINTS = ("获奖人数", "人数", "名额", "数量", "count")
HEADER_SCAN_ROWS = 10


@dataclass
class TierPreview:
    sheet: str
    tiers: list[PrizeTier]
    has_header: bool


def _find_column(headers: list[str], hints: tuple[str, ...]) -> int | None:
    for index, header in enumerate(headers):
        lowered = header.lower()
        if any(hint in lowered for hint in hints):
            return index
    return None


def _find_header_row(rows: list[list[str]]) -> int | None:
    """选择前若干行中最像表头的一行（命中最多类关键词）。

    这样可以跳过“2026 年会奖项设置”这类标题行。
    """
    best_index: int | None = None
    best_score = 0
    for index, row in enumerate(rows[:HEADER_SCAN_ROWS]):
        cells = [cell.lower() for cell in row]
        score = sum(
            1
            for hints in (TIER_NAME_HINTS, PRIZE_HINTS, COUNT_HINTS)
            if any(hint in cell for cell in cells for hint in hints)
        )
        if score > best_score:
            best_score = score
            best_index = index
    return best_index


def read_tiers(path: str | Path, sheet: str) -> TierPreview:
    """读取奖项表。

    表头自动识别（奖项/等级、奖品、人数/名额）。没有表头时按列顺序：
    3 列 = 奖项、奖品、人数；2 列 = 奖项、人数。
    """
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        if sheet not in wb.sheetnames:
            raise ValueError(f"工作表「{sheet}」不存在")
        rows = [[cell_text(v) for v in row] for row in wb[sheet].iter_rows(values_only=True)]
    finally:
        wb.close()

    while rows and not any(rows[-1]):
        rows.pop()
    if not rows:
        raise ValueError("该工作表没有数据")

    header_index = _find_header_row(rows)
    if header_index is not None:
        headers = rows[header_index]
        data_rows = rows[header_index + 1 :]
        has_header = True
    else:
        width = max(len(row) for row in rows)
        headers = [f"{get_column_letter(i + 1)} 列" for i in range(width)]
        data_rows = [row for row in rows if any(row)]
        has_header = False

    name_col = _find_column(headers, TIER_NAME_HINTS)
    prize_col = _find_column(headers, PRIZE_HINTS)
    count_col = _find_column(headers, COUNT_HINTS)

    if not has_header:
        width = len(headers)
        name_col = 0
        prize_col = 1 if width >= 3 else None
        count_col = width - 1 if width >= 2 else None

    if name_col is None:
        raise ValueError("未找到奖项名称列（表头需包含“奖项”或“等级”）")
    if count_col is None:
        raise ValueError("未找到获奖人数列（表头需包含“人数”或“名额”）")

    def cell(row: list[str], column: int | None) -> str:
        if column is None or column >= len(row):
            return ""
        return row[column]

    tiers: list[PrizeTier] = []
    for offset, row in enumerate(data_rows, start=1):
        name = cell(row, name_col).strip()
        if not name:
            continue
        count_text = cell(row, count_col).strip()
        if not count_text:
            raise ValueError(f"数据第 {offset} 行缺少获奖人数")
        match = re.search(r"\d+", count_text)
        if match is None:
            raise ValueError(f"数据第 {offset} 行的人数不是数字：{count_text}")
        try:
            tiers.append(
                PrizeTier(
                    name=name,
                    prize=cell(row, prize_col).strip(),
                    count=int(match.group()),
                )
            )
        except ValueError as exc:
            raise ValueError(f"数据第 {offset} 行：{exc}") from exc

    if not tiers:
        raise ValueError("未读取到任何奖项，请检查表格内容")
    return TierPreview(sheet=sheet, tiers=tiers, has_header=has_header)
