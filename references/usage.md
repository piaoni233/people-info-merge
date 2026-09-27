# 详细使用指南与命令参考（usage）

本文档提供 `scripts/people_merge.py` 完整命令行参数与各模式执行细节，供开发者、进阶用户与脚本调用方查阅。
基础概述与快速指引请见 `SKILL.md`。

---

## 一、命令概览

`scripts/people_merge.py` 提供 5 个子命令：

| 子命令 | 用途 | 适用时机 | 是否修改文件 |
|---|---|---|---|
| `inspect` | 探测输入文件元数据、表头与样本行 | 阶段一：初次接入新表 | 只读 |
| `suggest-keys` | 计算并推荐候选关联键，输出对齐率证据 | 阶段二：辅助用户确认 | 只读 |
| `merge` | 执行横向清洗、合并并生成 Excel 与统计日志 | 阶段三：执行合并任务 | 写入输出文件 |
| `selftest` | 运行内置两阶段端到端自测试（CSV 与跨格式） | 安装验证、环境排错 | 临时文件（自动清理） |
| `selfcheck` | AST 语法树安全扫描，检查高危模块与调用 | 上架前合规自检 | 只读 |

---

## 二、命令详细参数与使用方法

### 1. `inspect`：表头与样本行探测

只读扫描输入文件，提取字段名称、编码、嗅探格式与数据样本行（默认 5 行），不卷入全量数据。

```bash
python scripts/people_merge.py inspect \
  --files <file1> <file2> [file3...] \
  --sample-n 5 \
  --out profile.json
```

**参数说明：**
- `--files`（必需）：待合并的文件路径列表，支持混合格式（`.xlsx`、`.xlsm`、`.ods`、`.xls`、`.csv`、`.tsv`、`.txt`）。
- `--sample-n`（可选，默认 5）：每个表采样行数。
- `--out`（必需）：输出元数据 JSON 文件路径。

---

### 2. `suggest-keys`：关联键打分与重合度证据

基于 `inspect` 产出的 `profile.json`，评估各表列名相似度、值域重合度与唯一性，生成推荐键列表。

```bash
python scripts/people_merge.py suggest-keys \
  --profile profile.json \
  --out keys.json
```

**参数说明：**
- `--profile`（必需）：`inspect` 产出的元数据文件。
- `--out`（必需）：推荐结果与对齐证据 JSON 文件路径。

---

### 3. `merge`：执行横向合并

根据确认后的 `plan.json` 执行多表横向拼接。

```bash
python scripts/people_merge.py merge \
  --plan plan.json \
  --out 汇总结果.xlsx \
  --backend auto \
  --save-recipe recipe.json
```

**参数说明：**
- `--plan`（必需）：合并执行方案配置文件。
- `--out`（必需）：生成的汇总 Excel 文件路径（脚本会自动在同目录下生成同名的 `.report.json`）。
- `--backend`（可选，默认 `auto`）：
  - `auto`：检测到 `openpyxl` 则使用其加速，未安装则自动无感回退至标准库 `stdlib`。
  - `stdlib`：强制纯 Python 标准库后端（零外部依赖，极速轻量）。
  - `openpyxl`：强制使用 `openpyxl` 引擎（未安装时报错退出）。
- `--save-recipe`（可选）：将本次关联键与策略保存为配方文件，便于下次相同报表直接复用。
- `--recipe`（可选）：直接加载历史配方文件（跳过关联键推荐步骤）。

#### `plan.json` 配置规范示例：

```json
{
  "primary_file": "招聘系统导出.csv",
  "unmatched_strategy": "separate_sheet",
  "tables": [
    {
      "path": "examples/sample_data/招聘系统导出.csv",
      "key_col": "姓名"
    },
    {
      "path": "examples/sample_data/考勤系统导出.csv",
      "key_col": "员工姓名"
    },
    {
      "path": "examples/sample_data/培训平台导出.csv",
      "key_col": "姓名"
    }
  ]
}
```

- `primary_file`：主表基准文件名（主表列排列在结果表前部）。
- `unmatched_strategy`：未匹配记录策略，可选：
  - `separate_sheet`（推荐）：结果表保留并标记部分缺失，同时在单独的“未匹配记录”Sheet 完整列出。
  - `keep_marked`：仅在结果表打标保留。
  - `drop`：从结果表剔除，仅在“未匹配记录”Sheet 与日志中保留。
- `tables`：各文件对应的匹配列名（支持跨表列名不一致，如“姓名”对“员工姓名”）。

---

### 4. `selftest`：端到端自测试

自包含内置验证，不依赖任何外部测试文件：

```bash
python scripts/people_merge.py selftest --backend stdlib
```

- **阶段一**：构造内存 CSV 样本，验证经典三表清洗、对齐证据、三 Sheet 生成及写后读回。
- **阶段二**：动态构造 HTML 伪 `.xls`、OpenDocument `.ods` 和 Tab 分隔 `.tsv`，验证跨格式混拼端到端正确性。

---

### 5. `selfcheck`：安全合规静态审计

通过 Python 官方 `ast` 模块解析自身语法树，输出合规审计 JSON：

```bash
python scripts/people_merge.py selfcheck
```

- 审计断言：无网络模块（`urllib/requests/socket/http`）、无系统派生（`os.system/subprocess`）、无动态代码执行（`eval/exec`）。
