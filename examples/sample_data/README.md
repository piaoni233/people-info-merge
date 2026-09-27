# 示例数据说明（全部为虚构脱敏数据）

本目录的示例模拟 HR 典型多源汇总场景：

| 文件 | 模拟来源 | 关键列 | 说明 |
|---|---|---|---|
| `招聘系统导出.csv` | 招聘系统 | 姓名 + 手机 | 常规 CSV（UTF-8 BOM） |
| `考勤系统导出.csv` | 考勤系统 | 员工姓名 + 部门 + 工号 | 列名与前者不同，含全角空格 |
| `培训平台导出.csv` | 培训平台 | 姓名 + 课程 + 联系方式 | 同义列名（联系方式≈手机） |
| `旧系统导出.xls` | 老人事系统网页导出 | 姓名 + 手机 + 入职日期 | **扩展名是 `.xls`，内容其实是 HTML 表格**（真实系统常见），用于演示"按内容识别格式" |

所有姓名、手机号（`138000000xx` 号段）、部门、工号均为虚构，与真实个人无关。

## 故意植入的"脏数据"（用于演示 Skill 的核心价值）

| 序号 | 位置 | 现象 | 期望行为 |
|---|---|---|---|
| 1 | 招聘" 李四 " / 考勤"王　五" / 旧系统" 李四 " | 前后空格、全角空格 | 归一化后正常对上 |
| 2 | 考勤工号 `EMP-1001` | 跨系统前缀 | 按姓名合并时不受影响；按工号合并时自动去前缀 |
| 3 | 招聘赵六手机为空 | 空值 | 结果表留白，日志计数 |
| 4 | 招聘独有"陈十二"、考勤独有"孙八/冯十三"、培训独有"周九/褚十四" | 未匹配记录 | 进入"未匹配记录"sheet 并附最相近候选 |
| 5 | 李四手机：招聘 `13800000002` vs 培训 `13800000099` | 字段冲突 | 两列都保留 + 行标记"冲突" + 日志清单 |
| 6 | 李四入职日期：考勤 `2024-03-01` vs 旧系统 `2024-03-05` | 跨格式字段冲突 | 同上（示范"HTML 伪 xls 与 CSV 混用"的冲突暴露） |

## 快速演示（在技能目录下执行）

```bash
# ① 经典三表（CSV）
python scripts/people_merge.py inspect --files examples/sample_data/招聘系统导出.csv examples/sample_data/考勤系统导出.csv examples/sample_data/培训平台导出.csv --out profile.json
python scripts/people_merge.py suggest-keys --profile profile.json --out keys.json
# 按 keys.json 推荐确认关联键后，编写 plan.json（见 references/test_cases.md 用例 T1），再执行：
python scripts/people_merge.py merge --plan plan.json --out 汇总结果.xlsx --save-recipe recipe.json

# ② 混格式四表（CSV ×3 + 伪 .xls）
python scripts/people_merge.py inspect --files examples/sample_data/招聘系统导出.csv examples/sample_data/考勤系统导出.csv examples/sample_data/培训平台导出.csv examples/sample_data/旧系统导出.xls --out profile.json
# plan.json 内 4 个 table 的 key_col 依次为 姓名 / 员工姓名 / 姓名 / 姓名，再执行 merge
python scripts/people_merge.py merge --plan plan.json --out 混合格式汇总.xlsx
```

说明：`profile.json / keys.json / plan.json / 汇总结果.xlsx` 均为运行产物，
已被 `.gitignore` 忽略，不会进入上传包。
另：`.ods` 与真·二进制 `.xls` 的读取路径由 `python scripts/people_merge.py selftest`
与 `python -m unittest discover -s tests -v` 覆盖（测试用例按格式规范现场构造文件，无需外部样表）。

