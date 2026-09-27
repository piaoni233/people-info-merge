# 多格式表格解析矩阵（format_matrix）

本文档定义「多表人员信息汇总助手」针对各类办公表格格式的底层解析策略、特征嗅探原理、安全防护与依赖声明。

---

## 一、解析策略总览

本 Skill 采用 **内容/Magic Bytes 优先嗅探** 策略。输入文件的扩展名仅用于初步提示与合法性筛选，实际格式由文件头内容与特征签名动态确定：

| 格式分类 | 扩展名 | 真实特征判断 | 底层解析策略 | 必需依赖 | 安全机制与数据规整 |
|---|---|---|---|---|---|
| **现代工作簿** | `.xlsx` / `.xlsm` | PK zip 头，包含 `xl/workbook.xml` | 标准库 `zipfile` + `xml.etree.ElementTree`（可选 `openpyxl` 加速） | **零依赖** | 统一 `_safe_xml()` 阻断实体注入；对 `.xlsm` **绝不加载与执行宏代码**，只读单元格字符串 |
| **开放文档** | `.ods` | PK zip 头，包含 `content.xml` | 标准库 `zipfile` + XML 解析 | **零依赖** | 还原 `text:s` 空格与类型值；自动剔除 `number-columns-repeated` 展开后的末尾全空填充列 |
| **伪 .xls (HTML)** | `.xls` / `.html` | 文本/标签开头，含 `<table` 标签 | 标准库 `html.parser` 提取首个有效 `<table>` | **零依赖** | 展开 `colspan`；保留换行与转义字符；自动忽略 `<script>` 与 `<style>` |
| **伪 .xls (XML)** | `.xls` / `.xml` | XML 头，含 `urn:schemas-microsoft-com:office:spreadsheet` 命名空间 | 标准库 XML 解析（SpreadsheetML） | **零依赖** | 支持 `ss:Index` 跳列；数值类型自动去除多余 `.0`；自动丢弃全空填充列 |
| **定界文本** | `.csv` / `.tsv` / `.txt` | 纯文本数据 | 标准库 `csv.Sniffer` 自动探测定界符（`,`、`;`、`\t`、`\|`） | **零依赖** | 编码自动探测链：`utf-8-sig → utf-8 → gb18030 → gbk → big5` |
| **真·二进制 .xls** | `.xls` | OLE2/BIFF 复合二进制签名（`\xd0\xcf\x11\xe0`） | 声明可选依赖 `xlrd>=2.0.1`（BSD 开源） | **可选** | 未安装时绝不静默失败，提供“Excel/WPS 另存为 `.xlsx`（零依赖推荐）”与“`pip install xlrd`”明确指引 |

---

## 二、格式探测机制细节

1. **ZIP 包判定（.xlsx / .xlsm / .ods）**：
   - 先读取前 4 字节魔数 `PK\x03\x04`。
   - 若包内包含 `xl/workbook.xml`，判定为 Excel 现代工作簿（`.xlsx` 或 `.xlsm`）。
   - 若包内包含 `content.xml`，判定为 OpenDocument 工作表（`.ods`）。
   - 若非上述结构，抛出“非合规工作簿文件”拒绝解析。

2. **伪 .xls 探测**：
   - 相当一部分企业管理系统导出的 `.xls` 实际上是 HTML 表格或 Excel 2003 XML（SpreadsheetML）。
   - 引擎预读文件前 4096 字节文本：
     - 若包含 `<html` 或 `<table`，自动重定向至标准库 HTML 转换器解析；
     - 若包含 `urn:schemas-microsoft-com:office:spreadsheet`，自动重定向至 SpreadsheetML 解析器；
     - 若命中 BIFF 签名，且未安装 `xlrd`，明确指导用户另存为 `.xlsx` 格式。

3. **文本编码与分隔符自适应**：
   - 依次尝试 `utf-8-sig`、`utf-8`、`gb18030`、`gbk`、`big5` 逐级解码；
   - 提取前 2048 字符，交由 `csv.Sniffer` 判定分隔符；若探测失败，回退到按行频次统计规则（逗号/制表符/分号），确保老旧系统数据平滑读取。
