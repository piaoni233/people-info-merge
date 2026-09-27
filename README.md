# 多表人员信息汇总助手（People Info Merg）

![Python Version](https://img.shields.io/badge/python-3.9%2B-blue.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)
![Dependencies](<https://img.shields.io/badge/dependencies-Zero%20Required-brightgreen.svg>)
![Version](https://img.shields.io/badge/version-1.0.0-orange.svg)
![Author](https://img.shields.io/badge/author-ymyh-purple.svg)

> 面向 HR、行政、项目助理的多表人员信息合并工具，重点解决跨系统导出表格中“关联键不统一、同名列含义冲突、脏数据表面不一致、格式碎片化”等问题。

---

## 一、解决什么问题

### 1. 最容易出错的地方：关联键的不确定性

不同表格用不同字段关联同一个人。表 A 用“姓名”，表 B 用“工号”，表 C 虽有“姓名”列却混入了空格、全角字符或系统前缀（如 `EMP-1001` vs `1001`）。
手动 VLOOKUP 或盲目合并极易导致人员错位或漏对。本工具自动扫描列名，推荐候选关联键并展示清洗后的重合度证据，由业务人员确认后再执行合并。

### 2. 影响体验的地方：合并结果是否干净可用

合并后的结果如果列名混乱、空值未标记、未匹配记录被静默丢弃，用户依然需要花费大量时间二次整理。
本工具将所有列名打标来源归属、空值留白、未匹配记录单独输出至独立 Sheet，确保合并结果清晰、可核验。

### 3. 设计原则

- **本地计算**：纯本地环境执行，不依赖外部网络模块与云端模型；
- **只读元数据探测**：确认关联键前仅读取表头与样本行，不预先加载全量数据；
- **只在歧义处握手**：流程精简，仅设置两个必要确认点；
- **冲突不替用户选边**：遇到同一属性多表矛盾取值时，原样保留各表内容并标红暴露，由人工决断。

---

## 二、设计思路与架构决策

### 1. 解决什么问题

- **最容易出错的地方**：关联键的准确性。如果关联键找错或对错，后续所有字段合并都失去意义。因此，“关联键确认”是必须由用户拍板的强约束节点。
- **影响使用体验的地方**：结果报表的整洁度与未匹配去向。次要矛盾必须解决，但不能设置繁琐的多步阻断，因此采用默认策略推荐与一键确认。

### 2. 为什么确认点控制在两个？

参考办公 Agent 确认模式的分级设计原则，避免过度交互导致体验崩溃：

- **确认点一（关联键选定，中风险高歧义，必须）**：用户确认或修正关联列（如以“姓名”还是“工号”为基准）。未经确认不得执行合并。
- **确认点二（未匹配记录处置，中风险低歧义，建议）**：提供 `separate_sheet`（单独输出 Sheet）、`keep_marked`（结果表标记）与 `drop`（剔除）三个选项，默认推荐 `separate_sheet` 一键确认。

### 3. 为什么不做“全自动静默合并”？

自然语言或算法推断存在同名同姓、跨系统缩写等天然歧义。全自动看似省事，一旦静默匹配错误，人工倒查成本极高。通过展示“重合率百分比 + 独有人员示例”的人机握手机制，既省去配置负担，又守住数据准确底线。

### 4. 为什么选择纯标准库优先？

企业办公电脑与 Agent 运行环境往往受限于网络管控或缺失第三方包安装权限。通过 Python 3.9+ 标准库原生实现对 `.xlsx`（ZIP+XML）、`.ods`、HTML 伪 `.xls`、SpreadsheetML 与定界文本的解析，实现零安装开箱即用；同时无缝支持 `openpyxl`（读写加速）与 `xlrd`（旧版二进制 `.xls`）作为可选扩展。

---

## 三、项目工程结构

```text
多表人员信息汇总助手/
├── SKILL.md                     # Skill 入口元数据契约（面向平台与 Agent 执行指引）
├── README.md                    # 本文档（面向开发者与审核方：设计思路与工程说明）
├── LICENSE                      # MIT 开源许可证
├── .gitignore                   # 排除临时产物、运行输出与敏感缓存
├── scripts/
│   ├── people_merge.py          # 核心引擎：inspect / suggest-keys / merge / selftest / selfcheck
│   ├── requirements.txt         # 依赖说明（默认零必需依赖，openpyxl/xlrd 为可选依赖）
│   └── package_skill.py         # 自动化打包与合规校验脚本（保证规范打包，≤10MB）
├── references/                  # 规范证据链与详细参考文档
│   ├── usage.md                 # 完整命令行参数与配置规范
│   ├── output_format.md         # 三 Sheet 输出契约与列级定义
│   ├── troubleshooting.md       # 常见报错与排查速查表
│   ├── format_matrix.md         # 多格式底层解析矩阵与嗅探原理
│   ├── packaging.md             # 打包发布规范与合规验收清单
│   ├── privacy.md               # 数据隐私声明：纯本地内存处理
│   ├── security.md              # 安全审计：AST 扫描机制与实体注入防护
│   ├── third_party.md           # 第三方许可矩阵与架构决策
│   ├── test_cases.md            # 回归测试用例集（T1–T18 完整矩阵）
│   ├── limitations.md           # 风险边界与设计上限说明
│   └── changelog.md             # 版本变更记录
├── examples/sample_data/        # 虚构脱敏多源测试样本（化名与 138000000xx 号段）
│   ├── README.md                # 样本数据集说明
│   ├── 招聘系统导出.csv          # 招聘源（姓名 + 手机，常规 CSV）
│   ├── 考勤系统导出.csv          # 考勤源（员工姓名 + 部门 + 工号）
│   ├── 培训平台导出.csv          # 培训源（姓名 + 课程 + 联系方式）
│   └── 旧系统导出.xls           # 老系统导出（内容为 HTML 伪 .xls）
└── tests/                       # 自动化回归测试套件
    ├── test_backend_parity.py   # 标准库与 openpyxl 后端读写一致性测试
    └── test_people_merge.py     # 核心单元测试（覆盖全格式、拒收场景与业务逻辑）
```

---

## 四、快速开始

### 经典三步调用流程

1. **探测表头与样本元数据（只读扫描）**：

   ```bash
   python scripts/people_merge.py inspect \
     --files examples/sample_data/招聘系统导出.csv \
             examples/sample_data/考勤系统导出.csv \
     --out profile.json
   ```
2. **推荐候选关联键与重合度证据**：

   ```bash
   python scripts/people_merge.py suggest-keys \
     --profile profile.json \
     --out keys.json
   ```
3. **按确认后的方案执行合并**：

   ```bash
   python scripts/people_merge.py merge \
     --plan plan.json \
     --out 汇总结果.xlsx \
     --save-recipe recipe.json
   ```

*注：完整命令行参数、配置规范与配方复用方式详见 [references/usage.md](references/usage.md)；输出的三 Sheet 列级定义详见 [references/output_format.md](references/output_format.md)。*

---

## 五、开发、测试与打包

### 1. 运行安全自检与端到端测试

```bash
# AST 语法树静态安全审计（确认零网络导入、零外部进程派生）
python scripts/people_merge.py selfcheck

# 内置端到端自检（阶段 1：CSV 自检；阶段 2：HTML+ODS+TSV 跨格式合并）
python scripts/people_merge.py selftest --backend stdlib

# 执行全量单元测试套件（24 项回归断言）
python -m unittest discover -s tests -v
```

### 2. 构建交付包

```bash
python scripts/package_skill.py --out 多表人员信息汇总助手.zip
```

打包脚本会自动校验：

- `SKILL.md` 位于根层级路径；
- 过滤无关缓存与临时运行产物；
- 交付包体积远小于 10MB 上限。

---

## 六、架构决策记录（ADR）

- **为什么不使用 Pandas？**
  Pandas 体积庞大（包含 C 扩展通常超 100MB），且依赖复杂的动态编译环境。办公场景下数据量一般在 20 万行以内，Python 原生轻量数据结构即可在秒级完成处理，且便于在受限环境下极速分发。
- **为什么输出采用 3 个 Sheet 而非单个扁平表？**
  扁平表无法兼顾“主数据浏览”、“执行过程审计”与“异常排查”。拆分为 `汇总结果`、`合并日志` 与 `未匹配记录` 既满足直观消费，又保留了完整的责任追溯链条。
- **为什么不支持 PDF 与扫描件直接输入？**
  表格合并属于高精度确定性处理，OCR 提取包含不可控的错字与排版漂移风险。坚持“只处理结构化表格”，引导用户前置转换为 Excel/CSV，守住输出质量底线。

---

## 七、许可证与署名

- **许可证**：[MIT License](LICENSE)
- **作者**：`ymyh`
- **版本**：`1.0.0`
