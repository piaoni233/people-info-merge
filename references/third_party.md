# 第三方依赖与授权（third_party）

## 结论：无必需第三方依赖

- 合并引擎 `scripts/people_merge.py` 默认仅使用 **Python 3.9+ 标准库**：
  `argparse / ast / csv / datetime / difflib / html.parser / io / json / os / re / sys /
  tempfile / unicodedata / zipfile / xml.etree / xml.sax`。
- 表格解析（含 `.xlsx/.xlsm` 的 ZIP+XML、`.ods` 的 `content.xml`、HTML 表格、Excel 2003 XML、
  CSV/TSV 分隔符嗅探）全部由标准库完成，**无必需第三方依赖**。
- 标准库遵循 **PSF License 2.0**（随 Python 发行，可商用），无需额外授权文件。
- 本 Skill **不调用任何第三方 API、模型、SaaS 服务**，不存在数据出境：
  所有解析与合并都在用户设备侧完成（见 `privacy.md`）。

## 可选依赖（非必需、未捆绑）

| 组件 | 版本要求 | 来源 | 许可证 | 可商用 | 授权文件位置 | 数据出境 |
|---|---|---|---|---|---|---|
| openpyxl | >= 3.1.0 | PyPI（`pip install openpyxl`） | MIT | 是 | https://openpyxl.readthedocs.io/en/stable/ | 否（纯本地 .xlsx 读写，不联网） |
| xlrd | >= 2.0.1 | PyPI（`pip install xlrd`） | BSD | 是 | https://xlrd.readthedocs.io/en/latest/ | 否（纯本地 .xls 解析，不联网） |

- openpyxl 用途：`.xlsx` 读写加速/兜底。安装即用（`--backend auto/openpyxl`），
  未安装自动回退标准库后端（`--backend stdlib`），功能等价。
- xlrd 用途：**仅**读取旧版二进制 `.xls`（BIFF/OLE）。xlrd 2.x 官方说明
  “will no longer read anything other than .xls files”，因此它不参与 .xlsx/.xlsm 解析，
  权限面最小；未安装时对旧版 `.xls` 给出“另存为 .xlsx”或“pip install xlrd”的明确指引。
- 上传包内**不捆绑**上述依赖；`scripts/requirements.txt` 已按“名称、版本、许可证”登记备查。

## 决策记录（对应立项时的约定）

- 标准库后端经过 `selftest` + `tests/` 回归（含 .xlsx 写后读回校验、多格式读取用例）；
  若后续发现标准库路径 bug 数量超过可维护阈值，退路为“改用纯 openpyxl 方案”，
  届时将在此文档登记并发布新版本（见 `changelog.md`）。
- **不自行实现 BIFF 二进制解析**：真·二进制 `.xls` 只交给成熟的 xlrd（BSD），
  避免自制解析器带来误读风险；被改名成 `.xls` 的 HTML / Excel 2003 XML / 分隔符文本
  本身是文本格式，才由本 Skill 内置解析（见 `limitations.md`）。
