# 打包与发布规范（packaging）

本文档定义「多表人员信息汇总助手」在联想开放平台或其他 Agent 平台发布打包的规范、自动化校验机制与验收标准。

---

## 一、平台合规要求与交付规范

1. **入口契约位置**：
   - Skill 交付 ZIP 包必须以顶层目录为基准，保证解压后或根层级路径能够直接寻址到 `SKILL.md`。
   - 必须通过 `scripts/package_skill.py` 打包，统一路径分隔符为 `/`，适配各类 Linux/Windows 托管环境。

2. **体积与排除规范**：
   - 平台包体体积上限为 **10MB**。
   - 打包时必须严格排除：
     - 版本控制：`.git/`、`.svn/`；
     - 缓存与字节码：`__pycache__/`、`*.pyc`、`*.pyo`；
     - 运行产物与测试临时文件：`*.report.json`、`*.zip`、`plan.json`、`profile.json`、`recipe.json`、`汇总结果.xlsx` 等；
     - 本地敏感文件：`.env`、`*.pem`、`*.key`、`secrets/`。

---

## 二、自动化打包命令

在工程根目录运行：

```bash
python scripts/package_skill.py --out 多表人员信息汇总助手.zip
```

打包脚本内置多道前置断言检查：
1. 验证 `SKILL.md` 是否正确包含在打包清单中；
2. 验证生成的 ZIP 文件大小（远小于 10MB 上限）；
3. 校验 ZIP 归档内无任何隐藏的缓存目录或临时输出文件；
4. 输出最终的文件清单、文件总数与字节数。

---

## 三、发布前验收清单

在正式提交发布前，必须在本地依次执行并全部通过以下三项验证：

```bash
# 1. 静态安全扫描：验证零外部高危导入与子进程派生
python scripts/people_merge.py selfcheck

# 2. 内置端到端自测试：验证 CSV 与 HTML+ODS+TSV 跨格式合并
python scripts/people_merge.py selftest --backend stdlib

# 3. 完整回归测试套件：验证 24 项单元测试断言
python -m unittest discover -s tests -v
```
