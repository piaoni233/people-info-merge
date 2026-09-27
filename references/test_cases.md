# 测试用例（test_cases）

回归测试：`python -m unittest discover -s tests -v`（纯标准库，无需安装任何依赖）。
端到端自测：`python scripts/people_merge.py selftest --backend stdlib`。

## T1 正常三表合并（主路径，examples 数据）

- 输入：`examples/sample_data/` 三份 CSV。
- plan：三表键列分别为 `姓名 / 员工姓名 / 姓名`，主表招聘，策略 `separate_sheet`。
- 预期：全局 12 个不同人员；结果表 12 行；未匹配 9 条（赵六/钱七各 2 条来源行 + 陈十二/孙八/冯十三/周九/褚十四各 1 条）；
  3 个 sheet 齐全；`汇总结果.report.json` 落盘。5 个全命中人员：张三/李四/王五/吴十/郑十一。

## T2 关联键表面不一致

- 输入：招聘" 李四 "（前后空格）、考勤"王　五"（全角空格）。
- 预期：归一化后正常对上，不计入未匹配；`suggest-keys` 证据中 A/B 两侧对上率不受空格影响。

## T3 未匹配记录三种策略

- `separate_sheet`：未匹配行保留在结果表（标记"部分缺失"）+ 单独 sheet 备查。
- `keep_marked`：同上（结果表标记为唯一载体，sheet 仍保留）。
- `drop`：结果表仅保留全命中行，但日志中"未匹配记录条数"仍如实计数。

## T4 空值

- 输入：招聘赵六手机为空、培训钱七课程为空。
- 预期：结果表对应单元格留白；日志"空键值"仅统计键列，值列空值不报错。

## T5 表内重复键

- 输入：某表同一姓名出现两行。
- 预期：两行都被保留（取首行参与合并）；日志"表内重复键"列出组数与示例。

## T6 字段冲突暴露

- 输入：李四手机招聘 `13800000002` vs 培训 `13800000099`。
- 预期：两列都保留（`手机（来源：招聘系统导出.csv）` 与 `联系方式（来源：培训平台导出.csv）`，
  实测两列取值分别为 `13800000002` / `13800000099`）；该行末列 `数据状态` 为"冲突：联系方式"；
  日志"字段冲突清单"可查（实测 T1 数据冲突条数 = 1）。

## T7 多候选键歧义（姓名 vs 工号）

- 输入：考勤表同时有"员工姓名"与"工号（EMP- 前缀）"。
- 预期：`suggest-keys` 同时给出姓名配对与工号配对候选，按覆盖率排序；
  工号归一化去前缀后可比对（`EMP-1001`≡`1001`）。

## T8 GBK 编码 CSV

- 输入：GBK 编码的 CSV（无 BOM）。
- 预期：编码自动识别链兜底读出，不报"无法识别 CSV 编码"。

## T9 非法输入拒绝

- `.pdf` / 无扩展名 / 不存在路径 → `MergeError` 并给出转格式指引。
- plan 缺 `key_col` / 策略非法 / 键列不在表头 → `MergeError` 中文提示。
- 配方文件集合不一致 → "配方未命中，请重新确认关联键"。
- 配方命中（正例）：`--save-recipe` 落盘的配方再用 `--recipe` 加载，结果与人工确认路径
  逐格一致（用例 `test_recipe_reuse`；配方与 plan 之间按文件名比对，plan 无需写 `file` 字段）。

## T10 xlsx 写后读回（标准库后端自举）

- 用标准库后端写出含中文/空值/长文本的 3 sheet xlsx，再用标准库后端读回，
  行列数与关键单元格一致（selftest 内已断言）。

## T11 后端一致性（可选，需安装 openpyxl）

- `tests/test_backend_parity.py`：同一份 merge 输出分别用两后端读回，
  首 sheet 行列一致；未安装 openpyxl 时自动 skip。

## T12 结果表结构契约（实测断言）

- 首列必须为 `关联键`（归一化后的键值），末列必须为 `数据状态`（`完整匹配`/`部分缺失（缺：…）`/`冲突：…`）。
- 数据列名统一为 `原列名（来源：文件名）`，同一表的不同列按原表顺序排列，主表列在前。
- `tests/test_people_merge.py::TestMerge::test_merge_default` 已断言：李四行末列为"冲突：联系方式"，
  且手机相关两列都非空、取值不同（冲突不被静默合并）。

## T13 多格式读取（TSV / TXT 分隔符嗅探）

- 输入：`.tsv`（Tab 分隔）与 `.txt`（分号分隔）。
- 预期：`info["delimiter"]` 分别为 `\t` / `;`，表头与首行值正确，`info["format"] == "delimited"`。

## T14 伪 `.xls`（内容其实是 HTML 表格）

- 输入：`examples/sample_data/旧系统导出.xls`。
- 预期：`inspect` 中该表 `"format": "html"`，表头为 `姓名/手机/入职日期/来源系统`，8 行；
  带 `colspan` 的表头按展开处理（补齐空单元格）；`script/style` 与第二张 `<table>` 被忽略。
- 用例：`TestFormatReaders::test_html_disguised_xls`。

## T15 伪 `.xls`（内容其实是 Excel 2003 XML / SpreadsheetML）

- 输入：含 `xmlns='urn:schemas-microsoft-com:office:spreadsheet'` 的 XML，扩展名为 `.xls`。
- 预期：`info["format"] == "spreadsheetml"`；`ss:Index` 跳过的空列被丢弃；
  `ss:Type='Number'` 的 `1001` 不写成 `1001.0`。用例：`test_spreadsheetml_xls`。

## T16 `.ods`（OpenDocument）

- 输入：按 ODF 规范现场构造的最小 `.ods`（zip: `mimetype` + `content.xml`）。
- 预期：`info["format"] == "ods"`；表头 `姓名/联系方式/部门`（`number-columns-repeated='20'`
  的填充列被丢弃）；`office:value-type='float'` 的 `13800000001` 不带 `.0`；
  `text:s` 空格还原为 `王 五`；`office:value-type='date'` 取 `2024-03-01`。
  用例：`TestOdsReader::test_ods_read`（缺 `content.xml` 时须报错，见 `test_ods_without_content_xml`）。

## T17 真·二进制 `.xls` 与 XML 安全拒绝

- 输入：OLE 复合文档头 + 填充字节的 `.xls`。
- 预期：抛出 `MergeError`；未装 `xlrd` 时提示必须同时含“另存为 `.xlsx`”与“pip install xlrd”两条出路。
  用例：`TestRejects::test_xls_biff_guidance`（两种环境下都成立）。
- 输入：含 `<!DOCTYPE ... <!ENTITY ...>` 的 XML。
- 预期：`MergeError`（消息含“拒绝解析”），防实体膨胀。用例：`test_xml_with_entity_rejected`。

## T18 跨格式端到端合并（selftest 阶段 2）

- 输入：伪 `.xls`（HTML）+ `.ods` + `.tsv` 三份，按 `姓名` 合并。
- 预期：三表格式被识别为 `html / ods / delimited`；结果表 2 行（张三/李四）且无冲突；
  张三行同时含 `.tsv` 的“产品部”与 `.ods` 的“入职培训”。
