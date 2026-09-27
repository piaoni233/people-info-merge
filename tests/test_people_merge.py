#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""回归测试（纯标准库 unittest，强制 --backend stdlib / 直接调函数）。

运行：python -m unittest discover -s tests -v
"""
import json
import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "scripts")))
import people_merge as pm  # noqa: E402

BACKEND = "stdlib"
SAMPLE = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "examples", "sample_data"))
FILES = [os.path.join(SAMPLE, n) for n in ("招聘系统导出.csv", "考勤系统导出.csv", "培训平台导出.csv")]


def write_tmp(content: str, suffix=".csv") -> str:
    fd, p = tempfile.mkstemp(suffix=suffix)
    with os.fdopen(fd, "w", encoding="utf-8-sig", newline="") as f:
        f.write(content)
    return p


def write_tmp_bytes(data: bytes, suffix=".xlsx") -> str:
    fd, p = tempfile.mkstemp(suffix=suffix)
    with os.fdopen(fd, "wb") as f:
        f.write(data)
    return p


def unlink(*paths) -> None:
    for p in paths:
        try:
            os.remove(p)
        except OSError:
            pass


class TestNorm(unittest.TestCase):
    def test_spaces(self):
        self.assertEqual(pm.norm_key(" 李四 ", "name"), "李四")

    def test_fullwidth_space(self):
        self.assertEqual(pm.norm_key("王　五", "name"), "王五")

    def test_id_prefix(self):
        self.assertEqual(pm.norm_key("EMP-1001", "id"), "1001")
        self.assertEqual(pm.norm_key("工号：1002", "id"), "1002")

    def test_name_paren(self):
        self.assertEqual(pm.norm_key("张三（曾用名张小三）", "name"), "张三")


class TestCSV(unittest.TestCase):
    def test_read_sample(self):
        headers, data, info = pm.read_table(FILES[0], backend=BACKEND)
        self.assertIn("姓名", headers)
        self.assertEqual(info["n_rows"], 8)

    def test_reject_pdf(self):
        p = write_tmp("x", suffix=".pdf")
        try:
            with self.assertRaises(pm.MergeError):
                pm.read_table(p, backend=BACKEND)
        finally:
            os.remove(p)

    def test_reject_missing(self):
        with self.assertRaises(pm.MergeError):
            pm.read_table("不存在的文件.csv", backend=BACKEND)


class TestFormatReaders(unittest.TestCase):
    """多格式读取：TSV/TXT、伪 .xls（HTML / Excel 2003 XML）、.ods、旧版二进制 .xls。"""

    def test_tsv_and_semicolon_txt(self):
        p1 = write_tmp("姓名\t手机\n张三\t13800000001\n", suffix=".tsv")
        p2 = write_tmp("姓名;部门\n李四;设计部\n", suffix=".txt")
        try:
            h1, d1, i1 = pm.read_table(p1, backend=BACKEND)
            h2, d2, i2 = pm.read_table(p2, backend=BACKEND)
        finally:
            unlink(p1, p2)
        self.assertEqual(h1, ["姓名", "手机"])
        self.assertEqual(d1[0]["手机"], "13800000001")
        self.assertEqual(i1["format"], "delimited")
        self.assertEqual(i1["delimiter"], "\t")
        self.assertEqual(h2, ["姓名", "部门"])
        self.assertEqual(d2[0]["部门"], "设计部")
        self.assertEqual(i2["delimiter"], ";")

    def test_html_disguised_xls(self):
        html = (
            "<html><head><meta charset='utf-8'><style>td{color:red}</style></head>"
            "<body><p>员工花名册</p><table border='1'>"
            "<tr><th colspan='2'>姓名信息</th><th>部门</th></tr>"
            "<tr><td>张三</td><td>13800000001</td><td>产品部</td></tr>"
            "<tr><td> 李四 </td><td>&nbsp;</td><td>设计部</td></tr>"
            "</table><table><tr><td>第二张表应被忽略</td><td>x</td></tr></table>"
            "<script>var fake = \"<table><tr><td>脚本内的假表格</td></tr></table>\";</script>"
            "</body></html>")
        p = write_tmp(html, suffix=".xls")
        try:
            headers, data, info = pm.read_table(p, backend=BACKEND)
        finally:
            unlink(p)
        self.assertEqual(info["format"], "html")
        self.assertEqual(headers[0], "姓名信息")     # colspan=2 -> 后一列按空处理
        self.assertIn("部门", headers)
        self.assertEqual(len(data), 2)               # 只取第一张表
        self.assertEqual(data[0]["部门"], "产品部")
        self.assertEqual(data[1].get("姓名信息"), "李四")   # 去空格 / &nbsp; 归一
        self.assertEqual(len(headers), 3)

    def test_spreadsheetml_xls(self):
        xml = (
            "<?xml version='1.0'?><Workbook xmlns='urn:schemas-microsoft-com:office:spreadsheet'"
            " xmlns:ss='urn:schemas-microsoft-com:office:spreadsheet'>"
            "<Worksheet ss:Name='花名册'><Table>"
            "<Row><Cell><Data ss:Type='String'>姓名</Data></Cell>"
            "<Cell ss:Index='3'><Data ss:Type='String'>工号</Data></Cell></Row>"
            "<Row><Cell><Data ss:Type='String'>张三</Data></Cell>"
            "<Cell ss:Index='3'><Data ss:Type='Number'>1001</Data></Cell></Row>"
            "</Table></Worksheet></Workbook>")
        p = write_tmp(xml, suffix=".xls")
        try:
            headers, data, info = pm.read_table(p, backend=BACKEND)
        finally:
            unlink(p)
        self.assertEqual(info["format"], "spreadsheetml")
        self.assertEqual(headers, ["姓名", "工号"])   # ss:Index 跳过的空列被去掉
        self.assertEqual(data[0]["姓名"], "张三")
        self.assertEqual(data[0]["工号"], "1001")    # 数字不带 .0


class TestOdsReader(unittest.TestCase):
    """ODS 用例：测试时按 ODF 规范现场构造最小 .ods（zip + content.xml）。"""

    CONTENT = (
        "<?xml version='1.0' encoding='UTF-8'?>"
        "<office:document-content "
        "xmlns:office='urn:oasis:names:tc:opendocument:xmlns:office:1.0' "
        "xmlns:table='urn:oasis:names:tc:opendocument:xmlns:table:1.0' "
        "xmlns:text='urn:oasis:names:tc:opendocument:xmlns:text:1.0' office:version='1.2'>"
        "<office:body><office:spreadsheet>"
        "<table:table table:name='人员'>"
        "<table:table-row>"
        "<table:table-cell office:value-type='string'><text:p>姓名</text:p></table:table-cell>"
        "<table:table-cell office:value-type='string'><text:p>联系方式</text:p></table:table-cell>"
        "<table:table-cell office:value-type='string'><text:p>部门</text:p></table:table-cell>"
        "<table:table-cell table:number-columns-repeated='20'/>"
        "</table:table-row>"
        "<table:table-row>"
        "<table:table-cell office:value-type='string'><text:p>张三</text:p></table:table-cell>"
        "<table:table-cell office:value-type='float' office:value='13800000001'>"
        "<text:p>13800000001</text:p></table:table-cell>"
        "<table:table-cell office:value-type='string'><text:p>产品部</text:p></table:table-cell>"
        "</table:table-row>"
        "<table:table-row>"
        "<table:table-cell office:value-type='string'>"
        "<text:p>王<text:s text:c='1'/>五</text:p></table:table-cell>"
        "<table:table-cell office:value-type='string'/>"
        "<table:table-cell office:value-type='date' office:date-value='2024-03-01'/>"
        "</table:table-row>"
        "</table:table></office:spreadsheet></office:body></office:document-content>")

    def test_ods_read(self):
        import zipfile
        tmp = tempfile.mkdtemp(prefix="pm_ods_")
        p = os.path.join(tmp, "人员.ods")
        with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("mimetype", "application/vnd.oasis.opendocument.spreadsheet")
            zf.writestr("content.xml", self.CONTENT)
        try:
            headers, data, info = pm.read_table(p, backend=BACKEND)
        finally:
            import shutil
            shutil.rmtree(tmp, ignore_errors=True)
        self.assertEqual(info["format"], "ods")
        self.assertEqual(headers, ["姓名", "联系方式", "部门"])   # 稀疏填充列已去掉
        self.assertEqual(len(data), 2)
        self.assertEqual(data[0]["联系方式"], "13800000001")     # 浮点值不带 .0
        self.assertEqual(data[1]["姓名"], "王 五")                # text:s 空格还原
        self.assertEqual(data[1]["部门"], "2024-03-01")          # 日期取原始值

    def test_ods_without_content_xml(self):
        import zipfile
        tmp = tempfile.mkdtemp(prefix="pm_ods_")
        p = os.path.join(tmp, "坏文件.ods")
        with zipfile.ZipFile(p, "w") as zf:
            zf.writestr("styles.xml", "<x/>")
        try:
            with self.assertRaises(pm.MergeError):
                pm.read_table(p, backend=BACKEND)
        finally:
            import shutil
            shutil.rmtree(tmp, ignore_errors=True)


class TestRejects(unittest.TestCase):
    """拒收路径：真·二进制 .xls、含 XML 实体的文件、不支持的扩展名。"""

    def test_xls_biff_guidance(self):
        p = write_tmp_bytes(pm.OLE_MAGIC + b"\x00" * 1024, suffix=".xls")
        try:
            with self.assertRaises(pm.MergeError) as ctx:
                pm.read_table(p, backend=BACKEND)
        finally:
            unlink(p)
        msg = str(ctx.exception)
        # 无 xlrd 时应给出「另存为 .xlsx」+「pip install xlrd」两条出路；
        # 有 xlrd 时该伪造文件会解析失败，同样必须是 MergeError。
        self.assertTrue(("另存为" in msg and "xlrd" in msg) or "解析失败" in msg, msg)

    def test_xml_with_entity_rejected(self):
        xml = ("<?xml version='1.0'?><!DOCTYPE x [<!ENTITY a 'b'>]>"
               "<Workbook xmlns='urn:schemas-microsoft-com:office:spreadsheet'/>")
        p = write_tmp(xml, suffix=".xls")
        try:
            with self.assertRaises(pm.MergeError) as ctx:
                pm.read_table(p, backend=BACKEND)
        finally:
            unlink(p)
        self.assertIn("拒绝解析", str(ctx.exception))

    def test_unsupported_ext_hint(self):
        p = write_tmp("姓名,手机\n张三,13800000001\n", suffix=".pdf")
        try:
            with self.assertRaises(pm.MergeError) as ctx:
                pm.read_table(p, backend=BACKEND)
        finally:
            unlink(p)
        self.assertIn("改回正确扩展名", str(ctx.exception))

    def test_padding_column_dropped(self):
        """整列（含表头）都为空的填充列应被去掉。"""
        p = write_tmp("姓名,手机,\n张三,13800000001,\n李四,13800000002,\n")
        try:
            headers, _, _ = pm.read_table(p, backend=BACKEND)
        finally:
            unlink(p)
        self.assertEqual(headers, ["姓名", "手机"])


class TestKeys(unittest.TestCase):
    def test_top_candidate(self):
        prof = pm.cmd_inspect(FILES, sample_n=2)
        keys = pm.cmd_suggest_keys(prof)
        self.assertTrue(keys["candidates"])
        top = keys["candidates"][0]
        self.assertIn(top["a_kind"], ("name", "id"))
        self.assertGreaterEqual(top["score"], 0.5)
        self.assertIn("清洗后", top["evidence"])


class TestMerge(unittest.TestCase):
    def _plan(self, policy="separate_sheet"):
        return {"primary": os.path.basename(FILES[0]),
                "unmatched_policy": policy,
                "tables": [{"path": FILES[0], "key_col": "姓名"},
                           {"path": FILES[1], "key_col": "员工姓名"},
                           {"path": FILES[2], "key_col": "姓名"}]}

    def _run(self, policy):
        tmp = tempfile.mkdtemp(prefix="pm_test_")
        plan_p = os.path.join(tmp, "plan.json")
        out = os.path.join(tmp, "汇总结果.xlsx")
        with open(plan_p, "w", encoding="utf-8") as f:
            json.dump(self._plan(policy), f, ensure_ascii=False)
        try:
            rep = pm.cmd_merge(plan_p, out, backend=BACKEND)
            self.assertTrue(os.path.isfile(out))
            _, back, _ = pm.read_table(out, backend="stdlib")
            self.assertTrue(back)
            return rep, tmp
        except Exception:
            import shutil
            shutil.rmtree(tmp, ignore_errors=True)
            raise

    def test_merge_default(self):
        rep, tmp = self._run("separate_sheet")
        try:
            self.assertEqual(rep["n_keys"], 12)
            self.assertEqual(rep["n_unmatched"], 9)
            self.assertGreaterEqual(rep["n_conflicts"], 1)  # 李四手机号冲突
            self.assertEqual(rep["n_result_rows"], 12)
            # 冲突行:李四的数据状态含"冲突"，两列手机号都保留且值不同
            _, back, _ = pm.read_table(os.path.join(tmp, "汇总结果.xlsx"), backend="stdlib")
            lisi = [r for r in back if r.get("关联键") == "李四"][0]
            self.assertIn("冲突", lisi.get("数据状态", ""))
            phones = {k: v for k, v in lisi.items() if "手机" in k or "联系方式" in k}
            self.assertEqual(len([v for v in phones.values() if v]), 2)
            self.assertEqual(len(set(phones.values())), 2)
        finally:
            import shutil
            shutil.rmtree(tmp, ignore_errors=True)

    def test_merge_drop(self):
        rep, tmp = self._run("drop")
        try:
            self.assertLess(rep["n_result_rows"], 11)
            self.assertGreaterEqual(rep["n_unmatched"], 4)  # 日志仍计数
        finally:
            import shutil
            shutil.rmtree(tmp, ignore_errors=True)

    def test_plan_validation(self):
        tmp = tempfile.mkdtemp(prefix="pm_test_")
        bad = os.path.join(tmp, "plan.json")
        with open(bad, "w", encoding="utf-8") as f:
            json.dump({"tables": [{"path": FILES[0]}]}, f)
        try:
            with self.assertRaises(pm.MergeError):
                pm.cmd_merge(bad, os.path.join(tmp, "o.xlsx"), backend=BACKEND)
        finally:
            import shutil
            shutil.rmtree(tmp, ignore_errors=True)

    def test_recipe_mismatch(self):
        tmp = tempfile.mkdtemp(prefix="pm_test_")
        plan_p = os.path.join(tmp, "plan.json")
        rec_p = os.path.join(tmp, "recipe.json")
        with open(plan_p, "w", encoding="utf-8") as f:
            json.dump(self._plan(), f, ensure_ascii=False)
        with open(rec_p, "w", encoding="utf-8") as f:
            json.dump({"tables": [{"file": "别的表.csv", "key_col": "姓名"}]}, f)
        try:
            with self.assertRaises(pm.MergeError):
                pm.cmd_merge(plan_p, os.path.join(tmp, "o.xlsx"),
                             backend=BACKEND, recipe_path=rec_p)
        finally:
            import shutil
            shutil.rmtree(tmp, ignore_errors=True)

    def test_recipe_reuse(self):
        """先 --save-recipe 落盘配方，再用 --recipe 复用，结果应一致（回归：曾因 plan 无 file 键而误报未命中）。"""
        tmp = tempfile.mkdtemp(prefix="pm_test_")
        plan_p = os.path.join(tmp, "plan.json")
        rec_p = os.path.join(tmp, "recipe.json")
        with open(plan_p, "w", encoding="utf-8") as f:
            json.dump(self._plan(), f, ensure_ascii=False)
        try:
            rep1 = pm.cmd_merge(plan_p, os.path.join(tmp, "汇总结果.xlsx"),
                                backend=BACKEND, save_recipe=rec_p)
            self.assertTrue(os.path.isfile(rec_p))
            rep2 = pm.cmd_merge(plan_p, os.path.join(tmp, "汇总结果2.xlsx"),
                                backend=BACKEND, recipe_path=rec_p)
            self.assertEqual(rep1["n_result_rows"], rep2["n_result_rows"])
            self.assertEqual(rep1["n_unmatched"], rep2["n_unmatched"])
            self.assertTrue(rep2["recipe"], "配方复用应记录命中原因")
        finally:
            import shutil
            shutil.rmtree(tmp, ignore_errors=True)


class TestSelfcheck(unittest.TestCase):
    def test_ok(self):
        chk = pm.cmd_selfcheck()
        self.assertTrue(chk["ok"], chk)
        self.assertTrue(chk["backend_available"]["stdlib"])


if __name__ == "__main__":
    unittest.main()
