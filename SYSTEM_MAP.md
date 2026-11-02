# 宠速星系统地图与脚本登记表

本文件是新电脑、新 Codex 会话和维护人员接手本项目时的**统一入口**。它说明系统由什么组成、数据在哪里、每个脚本做什么，以及哪些脚本不能随意重跑。

## 接手时的阅读顺序

1. `AGENTS.md`：必须遵守的业务边界、财务口径和操作规则。
2. `SYSTEM_MAP.md`：系统结构、功能入口和脚本用途（本文件）。
3. `README.md`：新电脑恢复、授权和验证流程。
4. `agent/README.md`：飞书消息桥、定时同步和日常运维细节。

如果代码行为和文档不一致，先停止真实写入，查明差异并同步更新代码、测试和本文件，不能凭记忆执行。

## 数据与代码边界

| 内容 | 位置 | 是否进 Git | 说明 |
|---|---|---:|---|
| 程序、测试和说明 | 本仓库 | 是 | 换电脑后从 GitHub 恢复 |
| 飞书经营账、SKU 主表、驾驶舱 | 飞书云端 | 否 | 重新授权后连接原表，不从 Git 恢复 |
| 工作簿绑定 | `work/*_meta.json` | 否 | 新电脑由 `.example.json` 复制后填写 |
| 运行状态 | `D:\PetstarData\state` | 否 | 幂等、待确认、事务、诊断；需单独离线备份 |
| `agent/state` | 目录联接 | 否 | 指向 `D:\PetstarData\state` 的兼容入口 |
| 美团登录态和店铺绑定 | 本机 Chrome / 本地状态 | 否 | 只使用已验证的“宠速星宠物用品店”Chrome 资料 |
| token、二维码、Cookie | 本机授权目录 | 否 | 严禁提交或输出 |
| 导出成品 | `outputs/` | 否 | 本地交付物，可另行备份 |

## 主运行链路

```text
飞书消息
  -> agent/lark_event_bridge.ps1
  -> agent/codex_chat_bridge.py（理解和生成结构化动作）
  -> agent/record_finance_event.ps1（预览/确认/确定性执行）
  -> agent/finance_transaction.ps1（事务、回滚、幂等）
  -> 飞书经营表 + agent/state

Windows 计划任务
  -> agent/run_scheduled_job.ps1
  -> agent/run_sync_and_notify.ps1
  -> 美团只读提取 + 订单/成本/简账同步 + 完整性校验 + 飞书通知

每天 10:00
  -> agent/send_daily_traffic_suggestion.ps1
  -> 最新库存快照 + 最近 30 天自然订单（仅作参考）
  -> 飞书私聊一条“引流款运营建议”
```

同步采用三级路径：每小时只处理变化订单；每天 00:00 回看最近 7 天退款、核对全部有效补单，并只修相关订单/月份；每周一 00:00（周日晚结束后）先做全历史只读审计，只有审计失败或显式 `--full-rebuild` 才允许全量重建。证据不完整时停止写入并告警，不得自动扩大范围。普通补货只更新补货记录和 SKU 当前成本，不能自动倒改历史订单和简账。

指定订单或月份刷新若发现旧版数据把正数人工配送费只保存在月度收入表、而订单汇总仍为 0，`agent/sync_daily_ledger.ps1` 会在同一经营事务内按完整订单号把金额回写到订单汇总并增加“人工确认配送费”来源标记，然后继续当前操作。无法唯一定位订单时仍停止；不得静默丢弃或改写存在金额冲突的配送费。

## 当前功能登记

### 2026-11-02 经营费用统一来源

- 经营毛利 = 自然及线下订单扣佣后收入 - 货物/赠品成本 - 首次配送费；经营利润只减月份支出表经营费用一次，不另减干预订单总成本。
- `agent/expense_accounting.ps1` 为费用分类、待核对与月份汇总的共用入口；普通费用录入及简账同步只刷新受影响月份。月度结算 G 为经营费用合计，N 为推广费，O:R 分别为补单费用、售后追加配送费、其他经营费用、费用核对状态，均由支出明细生成。
- `work/update_dashboard_calculation_details.ps1` 部署口径并回填已有月份费用汇总，使用同一事务和回读验证；只改汇总与驾驶舱，不改历史订单、成本、退款或补货记录。
- 首次配送费仍写订单；售后追加配送费使用 OperatingExpense，保留原订单号，不覆盖首次配送费。自然/线下赠品并入订单；干预单实际赠品保留营销赠品支出。采购付款和老板取款不进入经营费用，未知费用显示待核对，重复的补单费用须告警。
- `work/test_expense_accounting.ps1` 离线测试分类、未知值、重复项和防重复扣费；禁止连接飞书。
- `agent/query_business_data.ps1`、`agent/run_sync_and_notify.ps1` 的经营利润查询/日结同用费用分类，`work/audit_finance_integrity.ps1` 只读比较支出明细与 G、N:Q 派生汇总，差异交由现有相关月份修复路径处理；不得因证据不完整自动扩大重建。`agent/codex_chat_bridge.py` 运行版本升为 `2026-11-02.1`，下次调用会自动加载售后费用新规则。

| 功能 | 主入口 | 读取/写入范围 | 使用说明 |
|---|---|---|---|
| 飞书对话记账 | `agent/lark_event_bridge.ps1`、`agent/codex_chat_bridge.py`、`agent/record_finance_event.ps1` | 读飞书消息；确认后写指定经营表 | 所有修改先预览，用户回复“确认”后原子执行；同一消息内多笔普通补货使用批量原子路径；店主明确纠正补货前成本账数量时，先执行 `SetCostBookQuantity` 再执行对应补货，整组原子确认，纠正值只作为移动平均成本旧数量，不修改美团库存、历史订单或简账；明确指定线下历史单修正 SKU 成本时，按只读查询返回的 `OFF-` 订单编号唯一定位，并原子重算该线下单成本/利润及所属月份；若用户给出的成本与 SKU 主表当前成本一致，则主表完全不动，只修目标线下单，不要求伪造平台订单号；同一待确认项先改 SKU 成本再登记线下单时，预演按前序新成本计算后续线下单；补充说明若没有形成可执行动作则保留原待确认计划，旧版遗留的零操作空计划会自动清除并按当前消息重建；Codex 在普通回答或待确认预览阶段返回 429 时，原消息都保持未完成并按 1—15 分钟递增冷却，由看门狗重放自动恢复，飞书只发送幂等的延后提示；Lark CLI 1.0.98 下公式区域拆开逐一校验 |
| 美团订单同步 | `work/sync_meituan_orders.ps1`、`work/meituan_order_extract.mjs`、`agent/order_item_identity_mappings.json` | 美团只读；写订单汇总/商品/干预数据 | 绑定唯一 Chrome 店铺；每次抓取及失败重试先回到第 1 页再完整分页；跨月日期范围先按自然月拆分读取，全部片段成功后才统一写账，避免月初最近 7 天退款回看被单月提取器拒绝；退款状态区分处理中/已完成/已关闭/无法识别，处理中或证据不足时不改账、保留观察结果并飞书提醒，后续同步继续追踪；店主确认“同商品其他规格未设置货号”后，可按商品名＋规格严格映射当前 KW/商品 ID，仅在平台货号为空时启用；公式校验使用 Lark CLI 1.0.98 支持的连续区域；每次按完整订单号用已存订单汇总/商品自愈历史干预单自动字段；不得修改美团订单或库存 |
| 美团库存同步 | `work/sync_meituan_inventory.ps1`、`work/meituan_inventory_export.mjs` | 美团只读导出；写库存相关飞书表 | 导出器默认字段保持稳定，附加字段通过环境变量请求 |
| 成本与利润 | `agent/rebuild_cost_snapshots.ps1` | 写订单成本快照和公式结果 | 历史快照默认冻结；只有明确指定订单才改历史成本；Lark CLI 1.0.98 下将多个公式区域拆开逐一校验 |
| 每日经营简账 | `agent/sync_daily_ledger.ps1` | 写月份收入/支出及驾驶舱来源数据；必要时把旧月表孤立的人工配送费原子归一到订单汇总 | 新单、退款、补单和赠品变化只刷新受影响月份；配送费来源分叉可安全自愈；补单平台扣款未知时留空并标记待识别；Lark CLI 1.0.98 下非连续公式区域逐一校验，禁止写成0或因校验参数格式回滚有效数据 |
| 推广费 | `agent/sync_promotion_expense.ps1`、`work/meituan_promotion_extract.mjs` | 美团只读；写经营支出 | 日结前读取当天累计消费；尚未结算时标记待补并自动重试，不阻断订单同步，不得猜金额 |
| 退款补偿 | `agent/refund_adjustments.ps1`、`agent/run_sync_and_notify.ps1` | 写已完成退款调整及对应成本；读取并提醒未完成/无法识别状态 | 按退款记录 ID 幂等；只有平台明确完成且金额大于 0 才调整收入；退款后商家预计收入归零时标记全额退款，否则标记部分退款；处理中或无法识别时不改收入、成本、库存并发送去重提醒；退回库存只冲回实际退回商品成本 |
| 实际履约修正 | `agent/record_finance_event.ps1`、`agent/order_actual_overrides.ps1` | 写指定订单的本地实际商品、最终收入和配送方式覆盖，再只刷新该单及所属月份 | 不修改美团库存；原商品数量不变且仅新增一个私下加购商品时，可按“最终收入－原商家收入”确定新增商品成交额；用户未提配送时确认执行沿用订单当前配送费；多个新增/替换商品仍须明确金额分摊 |
| SKU 主表维护 | `agent/sync_new_skus.ps1`、`agent/sync_sku_master_by_identity.ps1`、`agent/sku_master_schema.ps1` | 写 SKU 成本主表 | 写入前检查重复 KW 和商品身份冲突；普通补货遇到旧同号档案时，仅在“有商品ID且平台状态为售卖中”的候选恰好一个时定位该行，其他情况继续停止，且不会自动删除旧档案 |
| 赠品与干预单 | `agent/reconcile_pending_gifts.ps1`、`agent/intervention_record.ps1`、`agent/record_finance_event.ps1` | 写赠品、干预订单及其关联订单成本/月份 | 赠品成本计入订单成本，缺成本保持待确认；新增、撤销干预单及指定订单的履约/退款状态修改只调用订单级增量成本刷新，只写该订单、实际涉及的 SKU 行及所属月份，禁止为单笔确认重写 `SKU成本主表` 整列或刷新无关月份 |
| 经营数据查询 | `agent/query_business_data.ps1` | 只读飞书 | 供 Codex 回答经营、库存和待确认问题，禁止写入；SKU 查询返回绑定“主表行号＋美团商品ID”的内部 `action_code`，货号重复时必须用它安全定位，禁止裸商品ID或重复KW猜行 |
| 定时任务 | `agent/configure_petstar_scheduled_tasks.ps1` | 修改 Windows 计划任务（仅 `-Apply`） | 默认只预览；新电脑必须显式安装 |
| 驾驶舱月份选择 | `agent/update_dashboard_month_selector.ps1` | 写驾驶舱 `H1:I1` 月份选择、相关月度公式；`J` 列保存并隐藏系统当前月毛利 | 用户可切换历史月份；库存/同步状态仍显示当前值；定时通知始终读取隐藏的当前月毛利，不受用户选择影响；月份选项只在跨月时维护 |
| 每日引流款运营建议 | `agent/send_daily_traffic_suggestion.ps1`、`agent/traffic_suggestion_rules.json` | 只读最新库存快照和最近 30 天自然订单；仅写本地去重历史并向当前授权用户发飞书私聊 | 每天 10:00 固定推荐一个库存大于 0 且售卖中的商品；按关键词品类轮换，近期自然销量只显示为参考，不能作为是否推荐的门槛；发送正文与错误标签使用独立变量并有回归测试，人工修复补发使用独立幂等键；只建议标题、主图、折扣价、库存和推广状态等正常运营动作，不创建订单、不评价、不改库存 |
| 运行守护 | `agent/ensure_lark_event_bridge.ps1`、`agent/run_hidden_watchdog.vbs` | 管理本地进程；只读检查当前 Chrome/CDP 实例 | 每 5 分钟保持消息桥常驻、修复失联 CDP 代理；新 Chrome 实例只主动触发一次 Allow，授权恢复后自动补跑一次整点同步（夜间失败则补跑夜间恢复）；全程隐藏黑色窗口，不修改美团数据 |
| 美团在售 KW 折扣价导出 | `work/meituan_inventory_export.mjs`、`work/inspect_meituan_product_export_fields.mjs` | 美团商品页只读；本地生成 Excel | 导出时设置 `MEITUAN_EXPORT_EXTRA_FIELDS=price,discount_price_view`；折扣价为空才回退原价；多规格分行；非 KW 进入核对清单 |

### 商品折扣价导出规则

- `price` 是原价，`discount_price_view` 是当前折扣价。
- 只保留美团导出中“售中”的规格和 `KW` 开头货号。
- 一个货号多个规格或多个价格时必须分别列出，不能擅自合并价格。
- 折扣价缺失时使用原价，并在核对明细标明价格来源。
- `work/inspect_meituan_product_export_fields.mjs` 只读取导出弹窗字段并关闭弹窗，不创建导出、不改商品。
- 2026-10-21 首次交付文件位于本机 `outputs/meituan_price_export_20261021/`；`outputs/` 不进 Git。

## `agent/` 运行脚本登记

| 文件 | 用途 |
|---|---|
| `agent/AGENTS.md` | `agent/` 子目录的补充维护约束 |
| `agent/check_environment.ps1` | 新电脑或故障时只读检查依赖、路径和配置 |
| `agent/check_meituan_browser.mjs` | 有界检查当前 Chrome 实例与 CDP 代理；以 `DevToolsActivePort` 实例指纹和代理真实连接端口为准，不采信会跨重启残留的全局 Allow 标记；每个新实例未授权时主动连接并只触发一次 Allow；状态仅记录在 `agent/state/cdp_authorization_prompt.json`，不保存 Cookie/令牌；`--watchdog` 供五分钟看门狗短探测，`--mark-catch-up-started`/`--mark-sync-success` 只维护补跑幂等状态 |
| `agent/codex_chat_bridge.py` | 调用 Codex 理解飞书消息、查询数据并输出结构化动作；补货未明确赊账/未付款时默认当天全额付款；同条消息明确补货前成本账数量时输出 `SetCostBookQuantity` 并置于对应补货之前；遇到 429 时持久化全局冷却状态并返回可重试结果，成功调用后自动清除冷却；故障恢复可按原飞书消息ID精准失效旧解释缓存，不清空其他消息 |
| `agent/codex_response_schema.json` | Codex 动作协议的 JSON Schema |
| `agent/configure_hidden_gateway_startup.ps1` | 预览/配置 Hermes Gateway 隐藏启动 |
| `agent/configure_petstar_scheduled_tasks.ps1` | 预览/安装宠速星 Windows 计划任务 |
| `agent/send_daily_traffic_suggestion.ps1` | 每天 10:00 读取库存和近期自然销量，按关键词品类轮换生成一条简洁的引流款运营建议并私聊店主；同一天幂等，不连接美团、不修改经营表 |
| `agent/traffic_suggestion_rules.json` | 引流款关键词、标题匹配/排除词和合规运营动作配置；不包含订单创建、评价或销量干预动作 |
| `agent/update_dashboard_month_selector.ps1` | 幂等维护驾驶舱历史月份下拉和公式；默认保留用户当前选择，跨月只扩充选项；隐藏辅助列保存系统当前月毛利供定时通知使用 |
| `agent/ensure_lark_event_bridge.ps1` | 五分钟看门狗、消息重放、修复孤儿事件消费者，并提前检查 Chrome/CDP；代理死亡时后台重启，当前实例完成 Allow 后自动启动一次隐藏补跑；已有同步运行时不重复启动 |
| `agent/finance_transaction.ps1` | 财务写入互斥、before-image、提交与回滚 |
| `agent/hermes_chat_bridge.py` | 旧 Hermes 对话入口兼容层；新链路优先使用 Codex bridge |
| `agent/intervention_record.ps1` | 读取、匹配和维护干预订单记录 |
| `agent/lark_cli_noexit.ps1` | Lark CLI 包装，保留原退出码且不结束宿主进程 |
| `agent/lark_event_bridge.ps1` | 飞书事件长连接、去重、待确认队列和回执发送；普通成本账数量纠正与同条消息内补货合并预演、一次确认和原子执行；待确认补充未生成动作时不得覆盖原计划，历史零操作空计划会在下条消息进入时自动丢弃；Codex 普通调用或嵌套预览限流时均保留原消息供后台重放，并发送不泄露内部错误的幂等延后提示；所有回执返回值必须被抑制，禁止污染消息处理布尔状态 |
| `agent/order_actual_overrides.ps1` | 保存订单实际商品、配送方式等人工覆盖 |
| `agent/order_item_identity_mappings.json` | 保存店主已确认的“平台无货号规格→当前KW/美团商品ID”严格映射；只有商品名和规格同时匹配且平台货号为空时生效，不保存价格、成本、登录态或密钥 |
| `agent/petstar_cli.py` | 统一命令行入口和参数分发 |
| `agent/query_business_data.ps1` | 只读查询飞书经营数据 |
| `agent/README.md` | Agent 部署、运行、排障说明 |
| `agent/rebuild_cost_snapshots.ps1` | 按订单商品和成本规则重建/增量更新成本快照；逗号分隔的公式范围按连续区域逐一校验，避免 Lark CLI 1.0.98 拒绝非连续范围 |
| `agent/reconcile_pending_gifts.ps1` | 将已确认赠品成本补入相关订单 |
| `agent/record_finance_event.ps1` | 所有经营事件的确定性执行器；补货默认当天全额付款（明确赊账/未付款除外），且只更新补货记录与 SKU 当前成本；`SetCostBookQuantity` 仅纠正补货前成本账数量并保留原成本，和后续补货共用同一次主表读取及原子事务；内部 `sku_ref` 同时校验主表行号和商品ID，任一变化就停止，避免重复货号写错商品；指定 `OFF-` 线下订单的成本修正会回读该单售出商品并重算整单成本/利润，其他商品成本不明确时停止；新增/撤销干预单、指定订单实际履约和退款处置使用 `-Incremental -EligibleOrderNumbers <订单号>`，从订单汇总读取真实日期并只更新相关订单、SKU 与月份；若确认成本已等于主表基准成本和移动平均成本，则不改主表状态、备注或日期，只修目标线下单及其月份；公式校验将逗号分隔区域拆成连续范围，兼容 Lark CLI 1.0.98 |
| `agent/refund_adjustments.ps1` | 退款、退货和成本冲回的持久化规则 |
| `agent/run_hidden_gateway.vbs` | 无窗口启动 Gateway |
| `agent/run_hidden_scheduled_job.vbs` | 无窗口运行计划任务并返回真实退出码 |
| `agent/run_hidden_watchdog.vbs` | 无窗口运行消息桥看门狗 |
| `agent/run_scheduled_job.ps1` | 计划任务统一入口；除同步/日结外，分发每天 10:00 引流款运营建议，并在日结后按月维护驾驶舱月份选项；启动子任务前清除 Hermes/OpenClaw/Lark Channel 工作区路由标记，并固定使用当前 Windows 用户的 `~/.lark-cli` 授权目录，避免后台误读隔离工作区凭据；不复制或保存 OAuth token |
| `agent/authorize_lark_desktop.ps1` | 从普通 Windows 桌面/计划任务上下文发起一次飞书设备授权，并在同一上下文完成轮询和验证；仅在系统误把授权写入隔离注册表时人工运行，临时请求文件只含短期授权链接和 device code，不记录 token，禁止作为常驻任务 |
| `agent/run_sync_and_notify.ps1` | 分阶段同步、事务、审计和飞书通知主编排；每小时轻量同步，00:00 复核补单/近期退款并仅刷新受影响月份；平台退款处理中或识别不清时发送按状态去重的飞书提醒并继续追踪，不触发无关重算；Chrome 新实例未 Allow 时明确提示“实例授权”而非笼统代理故障，成功同步后清除待补跑标记；飞书凭证只读探针成功后再读取 SKU 表头，表头读取受统一重试、运行记录和失败通知保护，电脑离线后重新开机可在后续计划任务自动恢复；适配 Lark CLI 1.0.98，将逗号分隔的公式范围逐个按连续区域校验；成本快照子脚本失败时保留具体错误详情；安全诊断只记录 Windows Lark 凭据注册表是否存在及条目数，不读取或输出 token 内容 |
| `agent/run_weekly_maintenance.ps1` | 每周全历史只读审计；仅审计失败时触发原子化全量重建；不可作为普通同步入口 |
| `agent/runtime_paths.ps1` | 解析 Python、PowerShell、Lark CLI 和 web-access 等本机路径；服务账户取不到 Windows Shell 用户目录时回退到继承的 `USERPROFILE`，仍缺失则在外部读取/写入前停止 |
| `agent/sku_master_schema.ps1` | SKU 主表字段、规范化和安全访问函数 |
| `agent/start_hidden_gateway.ps1` | 隐藏启动 Gateway 的 PowerShell 包装 |
| `agent/sync_daily_ledger.ps1` | 同步月收入、月支出、格式和驾驶舱来源数据；在同一事务内自愈“旧月表有人工配送费但订单汇总为 0”的遗留分叉；公式校验逐个连续区域执行 |
| `agent/sync_new_skus.ps1` | 从平台发现并安全新增 SKU |
| `agent/sync_promotion_expense.ps1` | 读取并登记每日美团推广消费 |
| `agent/sync_sku_master_by_identity.ps1` | 按稳定商品身份合并/校正 SKU 主表 |

## `work/` 工具登记

### 当前维护与只读提取

| 文件 | 用途 |
|---|---|
| `work/audit_finance_integrity.ps1` | 审计订单、成本、月份简账、补单自动字段及经营数据一致性；支持指定月份的日结轻审计和无筛选的每周全历史审计 |
| `work/audit_sku_master.ps1` | 全面审计 SKU 主表字段和成本状态 |
| `work/audit_sku_master_duplicates.ps1` | 检查重复 KW、重复商品和冲突档案 |
| `work/inspect_meituan_product_export_fields.mjs` | 只读查看美团商品导出可选字段 |
| `work/meituan_browser_context.mjs` | 锁定并复用已验证的宠速星 Chrome 店铺上下文；当前绑定实例已在线但没有美团标签页时，直接在该实例中新建店铺页并再次验证店铺身份，不再依赖 `data:` 探针误判资料不可用 |
| `work/meituan_inventory_export.mjs` | 创建并下载美团商品/库存只读导出；支持附加字段 |
| `work/meituan_order_extract.mjs` | 从美团商家页只读提取订单、商品和退款信息；同时兼容普通退款 `refundInfoList` 与美团取退 `refundGoodsInfoList`；整单退货缺少独立金额时以订单顾客实付作为退款额并保留全量退货商品；退款记录保留处理中/已完成/已关闭/无法识别状态，不再过滤掉未完成售后；共享标签页无论停在哪一页，抓取前都会重置到第 1 页 |
| `work/test_meituan_refund_extraction.mjs` | 本地只读验证美团退款状态分类，覆盖申请中、成功、关闭、否定完成和未知文案 |
| `work/test_meituan_browser_recovery.mjs` | 使用临时目录和本地模拟端口验证 Chrome 新实例授权恢复：即使全局 Allow 标记残留，未真实连接的实例仍须判定待授权；真实连接后保留一次待补跑并可幂等领取。测试不访问美团、不写经营数据，临时文件结束即清理 |
| `work/meituan_promotion_extract.mjs` | 从美团推广页只读提取消费 |
| `work/reconcile_manual_delivery_fees.ps1` | 原子迁移/核对人工配送费权威来源 |
| `work/repair_dashboard_offline_formulas.ps1` | 修复驾驶舱线下订单口径并与简账核对 |
| `work/sync_meituan_inventory.ps1` | 把美团商品导出同步到飞书库存相关表 |
| `work/sync_meituan_orders.ps1` | 把美团只读提取结果同步到订单表；跨月区间按自然月拆读后合并，所有片段成功才进入写入；公式校验范围保持连续以兼容 Lark CLI 1.0.98 |

### 建表、初始化和通用维护

| 文件 | 用途 |
|---|---|
| `work/create_finance_workbook.ps1`、`work/create_simple_finance.ps1` | 创建早期/简版财务工作簿；简版驾驶舱经营毛利按“线上与线下收入-商品成本-配送费-平台推广费用”，经营利润再扣除不含推广费的其他经营费用；“计算明细”紧邻“数值”右侧，仅展示经营毛利、经营利润、现金净流入和期末现金的组成，库存货值按（美团库存+待恢复数量）×SKU 当前移动平均成本 |
| `work/update_dashboard_calculation_details.ps1` | 幂等更新现有驾驶舱的经营毛利/经营利润口径，把“计算明细”移动到“数值”右侧，并从当月支出回填月度结算“平台推广费用”；事务化、带公式校验和回读，可在口径被误改时重跑，不修改订单、库存或历史成本 |
| `work/ensure_replenishment_product_name.ps1` | 幂等确保 `补货记录` U 列为“商品名”；读取补货货号/商品 ID 与 `SKU成本主表` 的当前货号、商品 ID、自定义标题和平台标题，写入映射公式并校验回读；工作表仅有 A:T 时在 T 右侧追加 U 列，不移动既有列；事务化，可在列公式被破坏时重跑，不修改补货金额、成本、库存、订单或利润 |
| `work/create_daily_ledger.ps1` | 创建每日经营简账及月份表 |
| `work/fill_finance_formulas.ps1`、`work/write_finance_formulas.ps1` | 填充或重写财务公式 |
| `work/add_gift_feature.ps1` | 建立赠品字段和相关公式 |
| `work/add_sku_expiry_column.ps1` | 增加 SKU 到期日字段 |
| `work/import_local_cost_master.ps1` | 导入本地历史成本资料 |
| `work/initialize_dynamic_costs.ps1` | 初始化移动平均成本和批次相关数据 |
| `work/migrate_product_identity_and_batches.ps1` | 迁移稳定商品身份和成本批次结构 |
| `work/rebuild_meituan_history.ps1` | 明确需要时重建美团历史订单数据，属于重操作 |
| `work/update_intervention_details.ps1` | 更新干预订单明细 |
| `work/finance_guide.xml` | 财务工作簿说明/模板资源 |
| `work/*_meta.example.json` | 新电脑工作簿绑定模板，不含真实 token |

### 专项修复和历史一次性脚本

以下脚本保留用于审计追溯，**不得因文件存在就自动重跑**。执行前必须核对日期、订单、SKU、当前飞书状态，并先预览或备份：

- `apply_*`、`correct_*`、`fix_*`、`merge_*`、`normalize_*`、`repair_*`、`restore_*`：针对某次已知数据问题的专项改动。
- `audit_aug_sep_order_costs.ps1`、`build_cost_audit.py`、`analyze_local_costs.py`：历史成本分析和审计。
- `complete_recovered_confirmation.ps1`：完成特定恢复后的待确认状态。
- `inspect_intervention_sheet.ps1`、`diagnose_meituan_inventory.mjs`：只读诊断工具。
- `apply_missing_replenishments_20261003.ps1`、`apply_confirmed_costs_20260922.ps1`、所有文件名含日期或具体 KW 的脚本：一次性业务修复证据。
- `archive_and_cleanup_codex_backups.ps1`、`finalize_codex_backup_cleanup.ps1`、`migrate_codex_data_offline.ps1`、`run_codex_migration_offline.*`、`verify_codex_migration_offline.*`：Codex 数据迁移/归档工具；涉及文件移动和清理，必须另行明确确认。

逐文件用途如下：

| 文件 | 用途与限制 |
|---|---|
| `work/audit_aug_sep_order_costs.ps1` | 复核 2026 年 8—9 月历史订单成本，只读审计 |
| `work/analyze_local_costs.py`、`work/build_cost_audit.py` | 整理本地成本资料并生成审计中间结果，不是生产同步入口 |
| `work/apply_confirmed_costs_20260922.ps1` | 写入 2026-09-22 已人工确认的一批历史成本 |
| `work/apply_missing_replenishments_20261003.ps1` | 补录截至 2026-10-03 的指定漏记补货 |
| `work/complete_recovered_confirmation.ps1` | 在事务恢复后完成指定待确认项，必须核对 change ID |
| `work/configure_hidden_watchdog_task.ps1` | 旧版隐藏看门狗计划任务配置；新电脑优先使用 `agent/configure_petstar_scheduled_tasks.ps1` |
| `work/correct_kw4025_kw3058_costs_20261009.ps1` | 修正 KW4025、KW3058 的指定成本及目标订单 |
| `work/diagnose_meituan_inventory.mjs` | 只读诊断美团库存页结构并生成本地截图/状态 |
| `work/fix_kw1106_and_dental_costs_20260922.ps1` | 修正 KW1106 和牙膏牙刷历史成本分配 |
| `work/fix_kw2099_duplicates_and_gift_20261005.ps1` | 合并 KW2099 重复档案并修复指定赠品成本 |
| `work/fix_replenishment_date_formats.ps1` | 修复补货表日期显示格式 |
| `work/inspect_intervention_sheet.ps1` | 只读查看干预订单表结构和内容 |
| `work/merge_kw1119_cost_profile.ps1` | 合并 KW1119 的重复/分裂成本档案 |
| `work/merge_pending_sku_kw2131.ps1` | 将待确认 SKU 合并到 KW2131 |
| `work/normalize_gift_costs_20260922.ps1` | 统一指定历史赠品成本口径 |
| `work/repair_actual_receipt_20261008.ps1` | 修复 2026-10-08 指定实际收入记录 |
| `work/repair_duplicate_replenishment_kw2131.ps1` | 撤销/合并 KW2131 重复补货记录 |
| `work/repair_duplicate_sku_kw1113.ps1` | 合并 KW1113 重复 SKU 档案 |
| `work/repair_intervention_mapping.ps1` | 修复订单与干预记录的映射关系 |
| `work/repair_operating_expense_20260922.ps1` | 修复 2026-09-22 指定经营费用记录 |
| `work/repair_replenishment_kw1119_a884ceed.ps1` | 修复 change ID 为 A884CEED 的 KW1119 补货事务 |
| `work/repair_replenishment_production_dates.ps1` | 修复补货批次生产日期 |
| `work/repair_dashboard_and_replenishment_payments_20261025.ps1` | 一次性修复 2026-10-25 驾驶舱指标、库存成本货值公式及 2026 年 10 月缺失的补货付款；事务化回滚，不可盲目重跑 |
| `work/repair_sku_schema_shift_20260915.ps1` | 修复 2026-09-15 SKU 主表列错位事故 |
| `work/restore_lost_delivery_fees_20260914.ps1` | 恢复 2026-09-14 事故中丢失的配送费 |
| `work/archive_and_cleanup_codex_backups.ps1` | 归档并清理迁移后的 Codex 旧备份；具有清理动作，必须确认 |
| `work/finalize_codex_backup_cleanup.ps1` | 完成 Codex 旧备份清理；具有清理动作，必须确认 |
| `work/migrate_codex_data_offline.ps1` | Codex 数据离线迁移的主逻辑 |
| `work/run_codex_migration_offline.ps1`、`work/run_codex_migration_offline.cmd` | 关闭 Codex 后启动离线迁移 |
| `work/verify_codex_migration_offline.ps1`、`work/verify_codex_migration_offline.cmd` | 关闭 Codex 时校验迁移结果，不负责日常经营数据 |
| `work/simple_finance_meta.example.json` | 简版财务工作簿本机绑定模板 |
| `work/daily_ledger_meta.example.json` | 每日经营简账本机绑定模板 |

### 测试脚本

- `work/test_runtime_reliability.ps1`：核心静态契约、语法、路径和安全边界总测试；默认外部写入为 0。
- `work/test_codex_chat_bridge.py`、`work/test_hermes_chat_bridge.py`：消息理解桥测试。
- 其余 `work/test_*.ps1`：事务、工作簿、赠品、保质期、快速同步、SKU schema 等专项回归测试。
- `work/fake_lark_cli_for_transaction_test.ps1`：事务测试使用的假 Lark CLI，不能用于生产。

专项测试逐文件用途：

| 文件 | 验证内容 |
|---|---|
| `work/test_chunked_sheet_reads.ps1` | 飞书大范围读取拆分和完整性 |
| `work/test_expense_preservation.ps1` | 同步时人工支出不会丢失 |
| `work/test_expiry_notification.ps1` | 临期规则和通知 |
| `work/test_fast_sync_policy.ps1` | 无变化、普通新增和风险场景的轻重同步分流 |
| `work/test_finance_transaction.ps1` | 财务事务提交、失败和回滚 |
| `work/test_finance_transaction_storage.ps1` | 事务日志压缩、存储和恢复 |
| `work/test_finance_workbook.ps1` | 财务工作簿结构和核心公式 |
| `work/test_gift_order_formula.ps1` | 赠品成本计入订单利润 |
| `work/test_lark_event_bridge.ps1` | 飞书事件去重、待确认和重放 |
| `work/test_pending_gifts.ps1` | 未知赠品成本保持待确认 |
| `work/test_runtime_reliability.ps1` | 核心运行时契约；同时验证平台无货号规格映射必须严格、持久且不得覆盖平台已有货号 |
| `work/test_replenishment_product_name.ps1` | 补货商品名按商品 ID 优先、货号兜底，自定义标题优先及事务/公式安全性 |
| `work/test_sku_master_schema_safety.ps1` | SKU 主表字段访问和列变化安全性 |
| `work/test_sync_new_skus.ps1` | 新 SKU 发现、重复拦截和写入 |
| `work/test_trial_zero_cost_rule.ps1` | 试吃/尝鲜商品零成本规则 |

## 新增或修改功能时必须登记

每次新增或改变代码、脚本、计划任务、配置约定或用户可见功能，都必须在同一轮完成以下事项：

1. 更新本文件对应功能和脚本说明，写清主入口、读取内容、写入范围、依赖、触发方式和禁止事项。
2. 修改行为时同步更新 `AGENTS.md` 或 `agent/README.md`，不能只改代码。
3. 增加或更新自动化测试；至少运行 `work/test_runtime_reliability.ps1` 和相关专项测试。
4. 输出文件放 `outputs/`，密钥、登录态、二维码、运行状态和工作簿绑定不得提交。
5. 最终交付时询问是否提交；只有用户明确授权后才能 commit 和 push。

## 新电脑恢复完成标准

- 仓库代码与 GitHub 一致，`git status` 没有未知业务代码。
- 飞书 `user-default` 和机器人重新授权，工作簿绑定指向原表。
- Chrome 只绑定“宠速星宠物用品店”，只读预检通过。
- `D:\PetstarData\state` 已从离线备份恢复并正确联接。
- 环境检查、消息桥测试、可靠性测试全部通过。
- 计划任务经人工 `-Apply` 安装，手动跑一次同步且没有真实重复写入。
