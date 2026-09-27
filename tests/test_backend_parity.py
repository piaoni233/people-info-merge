#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""双后端一致性测试（可选）：需安装 openpyxl，否则自动 skip。

运行：python -m unittest discover -s tests -v
"""
import os
import sys
import tempfile
import unittest

try:
    import openpyxl  # noqa: F401
    HAVE = True
except Exception:
    HAVE = False

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "scripts")))
import people_merge as pm  # noqa: E402


@unittest.skipUnless(HAVE, "未安装 openpyxl，跳过后端一致性测试")
class TestBackendParity(unittest.TestCase):
    def test_write_read_parity(self):
        tmp = tempfile.mkdtemp(prefix="pm_parity_")
        try:
            sheets = [{"name": "汇总结果", "headers": ["关联键", "姓名", "备注"],
                       "rows": [["张三", "张三", ""], ["李四", " 李四", "空格已清洗"]]},
                      {"name": "合并日志", "headers": ["项目", "内容"],
                       "rows": [["执行时间", "2026-09-27"]]},
                      {"name": "未匹配记录", "headers": ["来源表", "原始键值"],
                       "rows": [["培训表", "周九"]]}]
            results = {}
            for be in ("stdlib", "openpyxl"):
                p = os.path.join(tmp, "out_%s.xlsx" % be)
                pm.write_result_xlsx(p, sheets, backend=be)
                _, back, _ = pm.read_table(p, backend=be)
                results[be] = (len(back), [r.get("关联键") for r in back])
            # 跨后端读回一致：行数与首列键值相同
            self.assertEqual(results["stdlib"], results["openpyxl"])
        finally:
            import shutil
            shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    unittest.main()
