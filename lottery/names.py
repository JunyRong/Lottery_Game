"""手动输入名单的解析。"""

from __future__ import annotations

import re

# 换行、英文/中文逗号、英文/中文分号、顿号、制表符
_SEPARATORS = re.compile(r"[\r\n,，;；、\t]+")


def parse_manual_names(text: str) -> list[str]:
    """把多行文本解析为姓名列表。

    - 支持每行一个姓名，也支持用逗号/分号/顿号/Tab 分隔；
    - 自动去掉首尾空白与空项；
    - 保留姓名内部空格（如 "John Doe"）和重复姓名。
    """
    return [part.strip() for part in _SEPARATORS.split(text) if part.strip()]
