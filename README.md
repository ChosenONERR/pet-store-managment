# 宠速星经营财务系统

本仓库保存可复用的程序、校验规则、测试和部署脚本。飞书经营表、运行状态、浏览器登录态、授权凭据和本机缓存不进入 Git；新电脑通过重新授权和恢复运行状态接回同一套业务数据。

## 仓库边界

可以提交：

- `agent/`：飞书消息桥、Codex 理解层、确定性写账与同步程序。
- `work/`：建表、审计、修复、测试和美团只读提取程序。
- `AGENTS.md`、说明文档、元数据示例和部署检查脚本。

不能提交：

- `agent/state/`、`agent/state.pre_migration_*/`：消息幂等、事务、日志、诊断和历史运行状态。
- 飞书 token、OAuth 状态、二维码、浏览器 Cookie、美团登录态。
- `work/simple_finance_meta.json`、`work/daily_ledger_meta.json`：当前飞书工作簿绑定。
- 临时导出、截图、审计结果和本地 npm/Python 缓存。

建议使用 GitHub 私有仓库。即使 `.gitignore` 已排除凭据和运行数据，代码仍包含本店业务规则与表结构。

## 新电脑恢复流程

1. 安装 Git、Node.js、Python、Codex CLI 和 Lark CLI。Hermes Agent 只在继续沿用旧调度器时需要；使用仓库内 Windows 计划任务后不再是必需依赖。
2. 克隆本仓库，不要把旧电脑的浏览器配置目录或 token 提交到 GitHub。
3. 在仓库根目录运行只读环境检查：

   ```powershell
   powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\agent\check_environment.ps1
   ```

4. 从示例创建本机工作簿绑定，并填写原飞书表格 URL：

   ```powershell
   Copy-Item .\work\simple_finance_meta.example.json .\work\simple_finance_meta.json
   Copy-Item .\work\daily_ledger_meta.example.json .\work\daily_ledger_meta.json
   ```

5. 重新完成 Lark CLI 的 `user-default` 用户授权和机器人授权。授权文件只留在新电脑。
6. 在 Chrome 登录美团商家后台，并启动 web-access 的本地 CDP 代理。登录态只留在新电脑。
7. 从单独的离线备份恢复 `D:\PetstarData\state`，再建立 `agent\state` 目录联接。没有旧状态也能重新运行，但会丢失本地消息幂等历史、待确认队列和诊断历史，因此不能称为“完全无缝迁移”。
8. 检查通过后，再注册五分钟看门狗和仓库内的隐藏定时任务。涉及 Windows 计划任务和启动项，必须在目标电脑上显式执行，不能由 Git clone 自动修改系统。`agent/configure_petstar_scheduled_tasks.ps1` 不带 `-Apply` 时只预览，不修改系统。

机器路径优先通过以下环境变量覆盖；不设置时脚本会自动查找常见安装位置：

- `PETSTAR_PYTHON`
- `PETSTAR_HERMES_PYTHON`
- `PETSTAR_LARK_CLI`
- `PETSTAR_WEB_ACCESS_CHECK`

## 黑色命令窗口

Codex 常驻 app-server 本身以后台子进程运行。桌面黑框来自 Hermes Gateway 使用控制台版 `python.exe` 启动，并被 Windows Terminal 接管。仓库内的 `agent/configure_hidden_gateway_startup.ps1` 可把登录启动入口切换到 `pythonw.exe` 包装器。更彻底的方案是先用 `agent/configure_petstar_scheduled_tasks.ps1 -Apply` 把经营定时任务迁到 Windows 计划任务，验证后再停用旧 Hermes Gateway；两步都属于系统配置，不能由仓库自动执行。

## 验证

代码修改后至少执行：

```powershell
python .\work\test_codex_chat_bridge.py
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\work\test_runtime_reliability.ps1
```

测试默认不得写入真实飞书经营数据。需要真实同步时，仍遵守预览、确认、原子写入、公式校验和失败回滚规则。
