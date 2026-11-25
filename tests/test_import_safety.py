"""上传文件名清洗的最小回归测试（标准库 unittest，无新增依赖）。

运行：python -m unittest tests.test_import_safety
"""
from __future__ import annotations

import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

from import_data import safe_filename


class TestSafeFilename(unittest.TestCase):
    def test_相对路径遍历被剥掉(self):
        self.assertEqual(safe_filename("../../evil.xlsx"), "evil.xlsx")

    def test_绝对路径只留文件名(self):
        self.assertEqual(safe_filename("/tmp/inbox/evil.xlsx"), "evil.xlsx")

    def test_纯文件名原样保留(self):
        self.assertEqual(safe_filename("1月红包发放台账.xlsx"), "1月红包发放台账.xlsx")

    def test_特殊名与空名回退为空(self):
        for bad in ("", ".", ".."):
            self.assertEqual(safe_filename(bad), "")

    def test_反斜杠路径永不逃出inbox(self):
        # Windows 下反斜杠是分隔符，会被剥掉；POSIX 下反斜杠不是分隔符，
        # 但不含 "/" 的名字无论怎样都只会落在 inbox 里，因此断言安全性质而非具体值
        out = safe_filename("..\\..\\evil.xlsx")
        self.assertNotIn("/", out)
        self.assertTrue(out.endswith(".xlsx"))


if __name__ == "__main__":
    unittest.main()
