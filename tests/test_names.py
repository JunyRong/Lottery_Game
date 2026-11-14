"""手动输入名单解析测试。"""

import unittest

from lottery.names import parse_manual_names


class TestParseManualNames(unittest.TestCase):
    def test_one_name_per_line(self):
        self.assertEqual(
            parse_manual_names("张三\n李四\n王五"), ["张三", "李四", "王五"]
        )

    def test_mixed_separators(self):
        text = "张三,李四，王五;赵六；钱七、孙八\t周九"
        self.assertEqual(
            parse_manual_names(text),
            ["张三", "李四", "王五", "赵六", "钱七", "孙八", "周九"],
        )

    def test_ignores_blank_lines_and_trims(self):
        text = "\n  张三  \n\n\t李四\t\n   \n"
        self.assertEqual(parse_manual_names(text), ["张三", "李四"])

    def test_keeps_internal_spaces(self):
        self.assertEqual(
            parse_manual_names("John Doe\nJane Roe"), ["John Doe", "Jane Roe"]
        )

    def test_keeps_duplicate_names(self):
        self.assertEqual(parse_manual_names("张三\n张三"), ["张三", "张三"])

    def test_only_separators(self):
        self.assertEqual(parse_manual_names(" ,;，、\n\t "), [])

    def test_windows_newlines_and_fullwidth_spaces(self):
        self.assertEqual(
            parse_manual_names("张三\r\n\u3000李四\u3000"), ["张三", "李四"]
        )


if __name__ == "__main__":
    unittest.main()
