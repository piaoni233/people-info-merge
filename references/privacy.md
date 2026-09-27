# 隐私声明（privacy）

## 1. 数据不出本地

- 本 Skill 的文件解析与合并逻辑（`scripts/people_merge.py`）**全部在用户设备侧执行**。
- 脚本不导入任何网络模块（`urllib / requests / socket / http / ssl` 等均未使用，
  可用 `python scripts/people_merge.py selfcheck` 自证），因此**不存在把原始表格
  内容上传到云端或外部服务的代码路径**。
- 联想开放平台的连接器能力与本 Skill 无关：本 Skill 不调用任何外部 API/模型。

## 2. 最小化读取

- `inspect` 阶段只读取**表头 + 前 N 行样本**（默认 5 行）生成 `profile.json`，
  不把全量数据卷入分析流程。
- 只有在用户确认关联键之后，`merge` 阶段才会读取全量行，且仅读写用户在
  命令行中明确指定的本地文件。

## 3. 示例数据全部虚构

- `examples/sample_data/` 中的姓名（张三/李四/王五……）、手机号（`138000000xx` 号段）、
  部门与课程均为**虚构脱敏数据**，与真实个人无关，仅用于演示与测试。
- `.gitignore` 已排除 `data/ input/ output/` 与 `*.xlsx`，防止真实业务表格被误提交。

## 4. 不落盘留存

- 合并过程不写任何缓存数据库；`selftest` 与单测全部在系统临时目录完成并自动清理。
- 输出只有用户指定的两样东西：`汇总结果.xlsx` 与 `汇总结果.report.json`（与 xlsx 同名后缀）。
- 建议用户：手机号等敏感列如需外发，请先自行脱敏；用后及时删除中间文件
  （`profile.json / keys.json / plan.json`）。
