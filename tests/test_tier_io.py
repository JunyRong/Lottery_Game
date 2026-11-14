"""奖项 Excel 导入测试。"""

import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from lottery.tier_io import read_tiers


def write_workbook(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "奖项"
    ws.append(["奖项名称", "奖品", "获奖人数"])
    ws.append(["一等奖", "笔记本电脑", 1])
    ws.append(["二等奖", "蓝牙耳机", "2人"])
    ws.append(["三等奖", None, 3])

    plain = wb.create_sheet("无表头")
    plain.append(["特等奖", "手机", 1])
    plain.append(["参与奖", "纪念品", 10])

    titled = wb.create_sheet("带说明")
    titled.append(["2026 年会奖项设置"])
    titled.append([])
    titled.append(["等级", "奖品", "名额"])
    titled.append(["金奖", "手表", 1])

    missing = wb.create_sheet("缺人数")
    missing.append(["奖项名称", "奖品"])
    missing.append(["一等奖", "电脑"])

    wb.create_sheet("空表")
    wb.save(path)


class TestTierImport(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.path = Path(self.tmp.name) / "奖项.xlsx"
        write_workbook(self.path)

    def tearDown(self):
        self.tmp.cleanup()

    def test_reads_header_sheet(self):
        preview = read_tiers(self.path, "奖项")
        self.assertTrue(preview.has_header)
        self.assertEqual([t.name for t in preview.tiers], ["一等奖", "二等奖", "三等奖"])
        self.assertEqual([t.prize for t in preview.tiers], ["笔记本电脑", "蓝牙耳机", "三等奖"])
        self.assertEqual([t.count for t in preview.tiers], [1, 2, 3])

    def test_reads_without_header(self):
        preview = read_tiers(self.path, "无表头")
        self.assertFalse(preview.has_header)
        self.assertEqual([(t.name, t.prize, t.count) for t in preview.tiers], [
            ("特等奖", "手机", 1),
            ("参与奖", "纪念品", 10),
        ])

    def test_header_after_title_and_blank_rows(self):
        preview = read_tiers(self.path, "带说明")
        self.assertTrue(preview.has_header)
        self.assertEqual(len(preview.tiers), 1)
        self.assertEqual(preview.tiers[0].name, "金奖")
        self.assertEqual(preview.tiers[0].count, 1)

    def test_missing_count_column_raises(self):
        with self.assertRaisesRegex(ValueError, "获奖人数"):
            read_tiers(self.path, "缺人数")

    def test_empty_sheet_raises(self):
        with self.assertRaises(ValueError):
            read_tiers(self.path, "空表")

    def test_unknown_sheet_raises(self):
        with self.assertRaises(ValueError):
            read_tiers(self.path, "不存在")


if __name__ == "__main__":
    unittest.main()
