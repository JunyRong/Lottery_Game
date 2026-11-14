"""Excel 导入/导出测试。"""

import random
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

from lottery.engine import LotteryEngine
from lottery.excel_io import (
    default_column,
    export_results,
    inspect_sheet,
    list_sheets,
    read_names,
)
from lottery.models import Entrant, PrizeTier


def write_workbook(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "抽奖名单"
    ws.append(["序号", "姓名", "部门"])
    ws.append([1, "张三", "研发"])
    ws.append([2, "李四", "市场"])
    ws.append([3, "张三", "销售"])
    ws.append([4, None, "财务"])
    plain = wb.create_sheet("无表头")
    plain.append(["王五"])
    plain.append(["赵六"])
    wb.create_sheet("空表")
    wb.save(path)


class TestExcelImport(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.path = Path(self.tmp.name) / "名单.xlsx"
        write_workbook(self.path)

    def tearDown(self):
        self.tmp.cleanup()

    def test_list_sheets(self):
        self.assertEqual(list_sheets(self.path), ["抽奖名单", "无表头", "空表"])

    def test_detects_header_and_name_column(self):
        preview = inspect_sheet(self.path, "抽奖名单")
        self.assertTrue(preview.has_header)
        self.assertEqual([c.header for c in preview.columns], ["序号", "姓名", "部门"])
        self.assertEqual(default_column(preview), 1)
        self.assertEqual(
            read_names(self.path, "抽奖名单", 1), ["张三", "李四", "张三"]
        )

    def test_without_header_uses_all_rows(self):
        preview = inspect_sheet(self.path, "无表头")
        self.assertFalse(preview.has_header)
        self.assertEqual(default_column(preview), 0)
        self.assertEqual(read_names(self.path, "无表头", 0), ["王五", "赵六"])

    def test_empty_sheet_has_no_columns(self):
        preview = inspect_sheet(self.path, "空表")
        self.assertEqual(preview.columns, [])
        self.assertIsNone(default_column(preview))

    def test_unknown_sheet_raises(self):
        with self.assertRaises(ValueError):
            inspect_sheet(self.path, "不存在")

    def test_skip_empty_cells_and_keep_duplicates(self):
        preview = inspect_sheet(self.path, "抽奖名单")
        names = preview.columns[1].values
        self.assertNotIn("", names)
        self.assertEqual(names.count("张三"), 2)


class TestExcelExport(unittest.TestCase):
    def test_export_contains_winners_and_summary(self):
        with tempfile.TemporaryDirectory() as tmp:
            out = Path(tmp) / "结果.xlsx"
            tiers = [
                PrizeTier(id="a", name="一等奖", prize="电脑", count=1),
                PrizeTier(id="b", name="二等奖", prize="耳机", count=2),
            ]
            engine = LotteryEngine(
                [Entrant(i, i + 1, name) for i, name in enumerate(["甲", "乙", "丙"])],
                tiers,
                rng=random.Random(3),
            )
            engine.draw("a")
            engine.draw("b")

            export_results(out, engine.winners, tiers)
            self.assertTrue(out.exists())

            wb = load_workbook(out)
            self.assertEqual(wb.sheetnames, ["中奖名单", "奖品统计"])
            rows = list(wb["中奖名单"].iter_rows(values_only=True))
            self.assertEqual(rows[0], ("编号", "奖项", "奖品", "姓名", "中奖时间"))
            self.assertEqual(len(rows), 3)
            self.assertEqual(
                rows[1][:4],
                (
                    engine.winners[0].number,
                    "一等奖",
                    "电脑",
                    engine.winners[0].name,
                ),
            )

            summary = list(wb["奖品统计"].iter_rows(values_only=True))
            self.assertEqual(summary[1][:5], ("一等奖", "电脑", 1, 1, "已完成"))
            self.assertEqual(summary[2][:5], ("二等奖", "耳机", 2, 1, "未完成"))


if __name__ == "__main__":
    unittest.main()
