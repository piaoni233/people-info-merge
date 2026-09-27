# 版本变更记录（changelog）

## 1.0.0 — 2026-09-27（首个正式对外发布版）

- **正式版发布**：
  - 功能闭环：关联键智能打分识别、交互确认机制、三 Sheet 标准输出（汇总结果 + 合并日志 + 未匹配记录）、配方保存与复用。
  - 健壮异常处理：全角半角空格规整、别名括号过滤、工号前缀剥离、同义字段冲突显式标记。
  - 主流办公格式全覆盖：原生支持 `.xlsx/.xlsm/.ods/.csv/.tsv/.txt` 及真实导出常见的各类改名伪 `.xls`（HTML / SpreadsheetML），真·二进制 `.xls` 友好指引。
  - 交付审计通过：纯本地零网络外发、XML 实体防攻击、单行表头消歧、24 项单元测试通过。

## 0.2.0 — 2026-09-27（预发布演进）

- **扩展支持办公全格式**：
  - 内容优先嗅探（`sniff_format`）：不再单纯信赖扩展名，支持真实 HR 导出场景中最普遍的“伪 .xls”。
  - **伪 `.xls`（HTML 表格导出）**：标准库 `HTMLParser` 透明解析首个 `<table>`，支持 `colspan` 展开、`<br>` 换行与表格嵌套，忽略 `script/style`。
  - **伪 `.xls`（Excel 2003 XML / SpreadsheetML）**：标准库 XML 解析，支持 `ss:Index` 跳列并剔除全空填充列。
  - **`.ods`（OpenDocument 电子表格）**：纯标准库（ZIP + `ElementTree`）解析，还原 `text:s` 空格、提取浮点/日期原始值，丢弃尾部全空填充列。
  - **`.xlsx` / `.xlsm`（宏工作簿）**：标准库 XML + 可选 openpyxl，仅提取单元格文本，**绝不执行 VBA 宏**。
  - **真·旧版二进制 `.xls`（BIFF/OLE 复合文档）**：增加可选依赖 `xlrd>=2.0.1`（BSD 许可）；未安装时明确提示两条出路（① 用 Excel/WPS 打开“另存为”`.xlsx`，推荐、零依赖；② `pip install xlrd`），绝不静默失败。
  - **`.csv` / `.tsv` / `.txt`**：自动识别 `utf-8-sig / utf-8 / gb18030 / gbk / big5` 编码与 `,`、`;`、`\t`、`|` 分隔符。
- **安全加固**：
  - 统一 XML 解析入口 `_safe_xml()`，遇到 `<!DOCTYPE` / `<!ENTITY` 直接拒绝解析，防止实体膨胀攻击（billion laughs）与 DTD 外部引用。
  - 单个 XML 解压上限限制为 64MB，重复列展开上限限制为 64。
- **自测与测试套件扩充**：
  - `selftest` 升级为两阶段，新增跨格式（HTML 伪 .xls + .ods + .tsv）端到端合并验证。
  - 新增 `TestFormatReaders`、`TestOdsReader`、`TestRejects` 等单元测试组，测试用例增加到 21 项（全部通过）。
  - `examples/sample_data/` 新增 `旧系统导出.xls`（HTML 表格实测样本）。

## 0.1.0 — 2026-09-27（首版）

- 新建 Skill 包：`SKILL.md`（入口契约）、`scripts/people_merge.py`（合并引擎）、
  `references/` 六件套、`examples/sample_data/` 虚构示例、`tests/` 回归测试。
- 合并引擎能力：`inspect`（只读表头+样本）/ `suggest-keys`（候选键+可读证据）/
  `merge`（汇总结果+合并日志+未匹配记录三 sheet 输出）/ `selftest` / `selfcheck`。
- 双后端：标准库后端（默认可用）+ openpyxl 可选后端（`--backend auto|stdlib|openpyxl`）；
  标准库后端经端到端自测，7 人三表场景：7 个不同人员、≥4 条未匹配。
- 结果表固定附带末列 `数据状态`（`完整匹配` / `部分缺失（缺：…）` / `冲突：…`）；
  字段冲突按"同义列名归一"识别（手机 / 联系电话 / 联系方式 → 同一逻辑字段），
  多表取值不一致时两列原值都保留、只做标记，不替用户判对错。
- 开发期修复（仍属 0.1.0 首版）：
  1. 结果表此前漏写 `数据状态` 列 → 现固定为末列（回归用例 `test_merge_default` 断言冲突标记可见）。
  2. `--recipe` 复用此前误报"配方解析失败：'file'"（plan 里只有 `path`，配方里是 `file`）
     → 现统一按文件名比对，新增回归用例 `test_recipe_reuse`。
  3. 打包脚本 `package_skill.py` 补 Windows 控制台 UTF-8 输出保护（与引擎一致，避免中文路径乱码）。
- 决策记录：若后续标准库 .xlsx 路径 bug 超过可维护阈值，退路为纯 openpyxl 方案，
  并在此文档与 `third_party.md` 中如实登记。
