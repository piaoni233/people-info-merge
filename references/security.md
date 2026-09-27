# 安全说明（security）

## 1. 无恶意代码

- `scripts/people_merge.py` 全部使用 Python 标准库（+可选的 openpyxl），
  不含混淆代码、下载器、持久化驻留、反调试等任何恶意行为。
- 可用 `python scripts/people_merge.py selfcheck` 自证：
  AST 级扫描确认**无网络模块导入**、**无外部命令调用**
  （`os.system / os.popen / os.exec* / subprocess.* / eval() / exec()` 均为零命中）。
- 也可以人工复核：`findstr /R "import urllib import requests import socket ^import http subprocess os\.system eval(" scripts\people_merge.py`
  应只命中本文件头部的文档注释与 `DANGER_PATTERNS` 自检元组（字符串字面量，非调用）。

## 2. 输入校验

- 扩展名白名单：仅 `.xlsx / .xlsm / .xls / .ods / .csv / .tsv / .txt`
  （`.pdf/.png/.doc/.docx`/无扩展名一律拒绝，并提示“若文件其实是表格，请改回正确扩展名”）。
  扩展名只是**入口提示**，真正的解析器由**内容嗅探**决定（ZIP 头 / OLE 头 / 文本特征），
  所以“把 xlsx 改名成 .csv”这类情况不会串到错误解析器上。
- 压缩包类（`.xlsx/.xlsm/.ods`）按结构校验：既无 `xl/workbook.xml` 也无 `content.xml`
  即判定“不是支持的表格格式”，ZIP 损坏（常见于加密文件）直接报错，
  **不尝试解密、不猜密码**。
- **XML 解析统一走 `_safe_xml()`**：含 `<!DOCTYPE` / `<!ENTITY` 的文件直接拒绝解析，
  防实体膨胀（billion laughs）与外部实体引用；`xml.etree` 本身不联网取 DTD。
- **不执行宏**：`.xlsm` 只按包内 XML 读单元格值，VBA 代码段从不加载/执行。
  真·二进制 `.xls` 交给可选的 xlrd（只读解析，不执行任何嵌入对象）。
- 解析时的压缩炸弹防护：压缩包内单个 XML 解压上限 64MB；单元格“重复列/行”展开上限 64；
  文本编码按 `utf-8-sig → utf-8 → gb18030 → gbk → big5` 依次尝试，全部失败则拒绝。
- 表头自动定位（首个含 ≥2 个非空单元格的行），空列名/重名列自动消歧为“列N”/“原名（2）”。
- 路径：只读写命令行明确指定的本地文件；输出路径的父目录不存在则由用户先创建，
  脚本不做目录遍历、不写计划外文件（除与 xlsx 同名的 `.report.json`）。

## 3. 资源上限（防超大文件拖死机器）

- 单文件 200MB 上限，单表 20 万行上限，超限直接拒绝并提示分批处理。
- xlsx 解析为一次性读入内存的轻量网格；列宽计算最多采样前 500 行。

## 4. 无密钥管理问题

- 本 Skill 不需要任何密钥、Token、账号：无登录、无云服务、无付费 API。
- `.gitignore` 已兜底忽略 `.env / *.key / *.pem / *.p12 / secrets/`，
  且源码中无任何硬编码凭据（selfcheck 的 AST 扫描可佐证调用面干净）。
