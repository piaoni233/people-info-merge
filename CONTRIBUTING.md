# 参与贡献指南（Contributing）

感谢你考虑为「多表人员信息汇总助手」做贡献。本项目面向 HR、行政、项目助理处理真实人员数据，**正确性优先于功能数量**：一次静默错配的代价远高于少一个功能。

## 贡献方式

- **报问题**：在 GitHub Issues 里描述现象、复现步骤、输入文件格式（请用 `examples/sample_data/` 的虚构样本或脱敏后的数据，**不要贴真实人员信息**）、期望结果与实际结果。
- **提 PR**：先 Fork 仓库，在自己的分支上改，提交 PR 到 `main`。PR 描述里写清改了什么、为什么、测试结果。一次 PR 只做一件事。

## 开发环境

- Python 3.9+，无需安装必需依赖（核心引擎只用标准库）。
- 可选依赖 `openpyxl >= 3.1.0`、`xlrd >= 2.0.1`（用于验证双后端一致性），没有也不影响开发与测试。

## 提交前必做的三项检查

```bash
# 1. 静态安全审计：确认零网络导入、零外部进程派生
python scripts/people_merge.py selfcheck

# 2. 内置端到端自测
python scripts/people_merge.py selftest --backend stdlib

# 3. 全量单元测试（24 项回归断言，须全部通过）
python -m unittest discover -s tests -v
```

## 代码规范

- **只用标准库**：核心引擎不得新增必需的第三方依赖；确有需要的功能先走可选依赖 + 自动回退。
- **不联网、不调外部模型**：引擎不得新增网络导入（`urllib`、`requests`、`socket`、`http`）、系统派生（`os.system`、`subprocess`）与动态执行（`eval`、`exec`），`selfcheck` 会拦截。
- **不替用户选边**：字段冲突时保留各表原值并标记，不静默取舍；未匹配记录不静默丢弃。
- **确定性优先**：同样的输入与方案必须产出同样的输出；不要引入随机性或依赖执行时序。
- **中文注释与文档**：注释、报错信息、文档都用中文；报错要给出明确的下一步（例如“请另存为 .xlsx”而不是只报异常）。

## 文档同步

- 改了命令参数、输出格式或行为边界，请同步更新 `references/` 下对应的文档（`usage.md`、`output_format.md`、`limitations.md`）与 `references/changelog.md`。
- 新增格式支持时，同步更新 `references/format_matrix.md` 与 `examples/sample_data/` 的样本。

## 许可证

本项目采用 MIT 许可证，提交的贡献默认同样遵循 MIT 许可证。
