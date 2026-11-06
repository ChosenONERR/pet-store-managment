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
  -> 飞书关键词库 + 最新库存快照 + 最近 30 天自然订单（仅作参考）
  -> 飞书私聊一条“引流款运营建议”
```

同步采用三级路径：每小时只处理变化订单；每天 00:00 回看最近 7 天退款、核对全部有效补单，并只修相关订单/月份；每周一 00:00（周日晚结束后）先做全历史只读审计，只有审计失败或显式 `--full-rebuild` 才允许全量重建。证据不完整时停止写入并告警，不得自动扩大范围。普通补货只更新补货记录和 SKU 当前成本，不能自动倒改历史订单和简账。

指定订单或月份刷新若发现旧版数据把正数人工配送费只保存在月度收入表、而订单汇总仍为 0，`agent/sync_daily_ledger.ps1` 会在同一经营事务内按完整订单号把金额回写到订单汇总并增加“人工确认配送费”来源标记，然后继续当前操作。无法唯一定位订单时仍停止；不得静默丢弃或改写存在金额冲突的配送费。

## 当前功能登记

### 2026-11-06 后台授权恢复与夜间完成状态修复

- `agent/check_lark_user_access.ps1` 为临时Windows只读诊断任务入口：固定当前用户授权目录，读取驾驶舱A1单元格，再读取本地授权状态；只保存安全分类和运行环境，不登录、不复制或输出凭据，不写经营表。临时任务须经店主同意创建，清理另行确认。
- 例行同步以实际最小读表验证用户权限，状态查询仅取本地元数据，不先执行额外的 `auth status --verify` 网络校验。读取失败仍停止经营写入；不得把维护窗口成功当作真实计划任务成功。
- 夜间恢复区分“业务待核对”和“技术步骤未完成”：前者保留飞书提醒但不触发整轮重跑；推广未取得、库存降级、配送查询失败等技术问题仍允许恢复。独立完成记录不被白天小时同步覆盖；正式完成与失败边界须经离线测试和真实后台验证。
- `agent/runtime_paths.ps1` 的共用完成策略由 `run_scheduled_job.ps1`、`ensure_lark_event_bridge.ps1` 同时使用；`run_sync_and_notify.ps1` 只为成功、非预演、未跳过订单/库存且无技术缺口的夜间结果写 `last_completed_night.json`。保留旧健康记录兼容，不用小时结果或人工核对警告推断夜间失败。看门狗还识别正在执行的原生PowerShell同步，避免重复启动。
- 2026-11-06 13:01真实临时任务 `Petstar-Lark-Diagnostic-20261106` 以小时任务相同Windows用户/交互会话运行，最小读表退出码3、用户状态missing、注册表可见条目2个，确认后台尚未恢复。此任务只按店主授权做诊断，未重新登录、未复制token、未改经营数据；等待店主再次同意真实后台授权及本轮临时任务清理，禁止自动删除其他旧授权任务。
- 本地CLI日志可确认01:52的一次 `auth status --verify` 调用刷新接口得到HTTP400，但未留存服务端业务错误码，不能断言是账户切换、普通过期或某个特定平台错误。官方v1.0.98 `internal/auth/uat_client.go` 的不可重试错误分支会清除该代失效凭据；这解释了可能的missing路径，不等于已证明具体拒绝原因。移除重复联网校验可减少无谓刷新，不能承诺平台永不撤销授权。
- 验证：23项授权顺序/夜间完成边界专项、24项全天调度、15项夜间恢复、65个PowerShell入口及消息桥回归通过，真实账表写入0。新增专项为 `work/test_lark_night_completion.ps1`，使用内存响应与临时状态文件，不读取真实凭据；真实小时同步和新一轮夜间完成记录仍需授权后验证，当前不报告恢复成功。
- 店主本轮已明确同意重新发起后台飞书授权，并允许验证结束后清理 `Petstar-Lark-Diagnostic-20261106`。临时授权任务 `Petstar-Lark-Reauthorize-20261106` 复用现有桌面授权脚本和小时任务同一交互用户，仅在本机临时目录保存短期申请及安全结果；打开官方确认页后等待店主点击，再做后台只读探针及正式小时同步验证。不读取或复用11月5日过期申请，不复制维护窗口凭据，不修改原定时任务计划。
- 本次13:35已打开官方确认页，后台授权轮询13:35:42提前退出（安全分类 `authorization_command_failed`、退出码3），不能报告授权成功。当前申请有效期600秒；若店主在有效期内完成确认，使用同一临时任务及申请 `-Resume` 续接，不盲目重复申请。新任务和诊断任务保留至验证结束；本次仅诊断任务已获清理许可，其他旧任务不得删除。
- 15:22店主告知已授权后，实际任务记录显示14:08:27及15:02:24小时同步已成功（非预演、非跳过、真实最小读表通过），故不再续接13:35的过期申请。临时诊断脚本15:23出现CLI退出码0但解析阶段报错，修正范围仅安全诊断封装，不因该误报重新授权或清除凭据；修复后再回读后台探针，完成后仅清理已获准的诊断任务。
- 临时诊断的误报已最小复现为Windows PowerShell后台默认编码与CLI的UTF-8输出不一致；补齐输出编码设置后，15:30:17实际后台探针成功（读表退出码0、user ready/valid），并加入离线防回归。该问题仅在本轮新增诊断工具中，不能据此否定14/15点正式同步，也不能将其误作凌晨授权missing的根因。等待手动启动正式小时任务复验，夜间完成策略仍只有离线验收，尚未等待下一次午夜。
- 最终验收：15:31:53主动启动原 `Petstar-Hourly-Sync`，15:34:21完成、任务退出码0，非预演、非跳过，实际用户读表通过，当天1单及配送查询完成、配送待查0单，无变化轻量路径未重算成本/简账。下一次16:00，全天触发器未改。唯一业务警告仍为历史补单 `3402374350443155017` 的配送确认字段缺失，未据此猜费用或擅改订单；结果 `recovery_required=false`，不触发技术性整轮补跑。
- 27项授权/完成记录专项（含真实Save-RunState临时状态写入及编码防回归）、24项调度、15项夜间恢复、65个PowerShell入口契约通过；离线测试不访问真实凭据和经营表。夜间完成策略尚未经历下一次真实午夜，只声明离线验证完成。
- 按店主明确许可，仅清理本轮 `Petstar-Lark-Diagnostic-20261106` 临时任务；脚本及安全诊断结果保留可供重新创建。未删除临时授权任务、11月5日旧任务、授权文件或任何凭据。
- 发布边界：店主随后明确确认将本轮改动提交并推送GitHub；仅纳入上述程序、离线测试和说明，不上传运行状态、授权申请、日志、工作簿绑定或任何凭据。授权拒绝的具体业务错误码缺失，无法证明是账户切换或普通令牌到期；减少重复联网检查不能替代平台最终授权，也不能保证永不需要人工重新授权。

### 2026-11-05 私下退款与同货号多规格实际履约修正

- `agent/record_finance_event.ps1` 的 `SetOrderActuals` 优先使用该订单商品行的稳定商品ID定位成本主表，不再仅凭重复KW拒绝已明确规格的订单；显式SKU定位符仍需核对主表身份。同一订单同号多规格、已知原ID无法唯一匹配或主表定位失效时停止，不按相同成本猜规格；缺少原ID的旧记录仅沿用原有唯一货号匹配，不放宽重复货号限制。
- 调整范围为店主明确指定订单的最终实收、实际交付商品、该单成本账及所属月份简账；保留未指定的原配送费，不改SKU单价、美团商品库存或其他历史订单。不新增一笔退款经营支出，避免收入调减后重复扣款。
- 实际履约结果补齐标准 `date` 字段，确保后续简账只刷新订单所属月份；执行器失败同时输出结构化业务原因，避免原生stderr只留下脚本行号。
- `work/test_order_actual_sku_identity.ps1` 以内存主表和原订单验证身份优先、冲突拒绝、零/缺失成本边界及预览路径；不得向飞书写测试数据。真实补记使用既有事务入口，先预览后执行并回读公式及订单利润；原失败消息仅在成功补记后停止重试。
- 已完成验收：23项专项、23项赠品、8项补货、33项轻量策略、24项调度、54项配送及消息桥/63个PowerShell入口运行检查通过，正式无效规格预览也能透传业务原因且零写入。原失败请求由后台恢复为 `A5446878`，店主确认后使用既有RetryPending执行；事务 `confirm-A5446878-20261105112104` 于11:29:15提交，成功回执11:29:18，待确认队列清空。16:05再次回读11月4日美团3号单 `3402373820497768164`，KW4025按商品ID51180062747实际交付5条、KW4152保留1瓶，最终实收11.96元、货物成本9.44元、配送5.60元、利润-3.08元；订单汇总与11月收入第16行一致，后续小时同步未覆盖人工履约修正。不改平台库存，不额外重复登记退款支出，不可把此历史补记当作新退款重复扣款。

### 2026-11-05 全天每小时同步

- `agent/configure_petstar_scheduled_tasks.ps1` 将普通小时触发器扩展为01:00—22:00，23:00复用原补录提醒、00:00复用日结入口，覆盖全天24个小时而不设置同刻重复同步。新增 `-HourlyOnly`，默认预览，`-Apply -HourlyOnly` 仅更新既有小时任务触发器并回读验证；不重装看门狗、提醒、推广、周审计或授权任务，不修改账户/登录方式/密钥。
- `agent/run_scheduled_job.ps1` 的Night/Weekly/NightRecovery在成功完成前一日日结后，串行调用当天轻量同步（SkipInventory、Silent），周一先完成该轻量同步再执行周审计，避免周审计延误凌晨新订单；23点原Reminder不变。失败保持真实退出码，不带半成品继续写入，不把小时同步升级成全量重建。
- 00:20/00:40恢复改读 `last_healthy_night.json`，当前日轻量同步覆盖 `last_run.json` 不会让已成功日结重复重建。`work/test_all_day_schedule.ps1` 通过假任务接口和实际调度分支离线验证24小时覆盖、跨月日期、执行次序、失败中断及健康日结判断，外部经营写入为0。
- 运行仍依赖本机开机、Chrome已授权且平台已登录；关机时本地任务无法执行。配送未结算继续待查，补录后仅刷新相关订单/月份，不能承诺00点拿到未来才结算的骑手费用。
- 安装验收：2026-11-05仅更新小时任务的22个触发器，原账户、执行入口和设置保持不变，其他9个任务XML对比无变化。10:23:45实际启动后台小时任务，10:26:12同步成功，任务退出码0、下次11:00。24项离线调度、15项夜间恢复、25项PowerShell及54项Node配送、23项赠品、8项补货、33项轻量策略和消息桥回归通过；62个PowerShell入口可靠性检查通过。午夜复合入口的执行顺序已离线验证，尚未等待真实午夜触发。

### 2026-11-05 夜间事务恢复与跨日配送费复查

- `work/recover_night_transaction_20261105.ps1` 仅恢复 `TX-20261105-000009-CD6BC333` 的写前快照，默认预览，`-Apply` 才在同步锁及经营写锁内回滚；不恢复其他事务、不改平台数据、补货或历史成本。必须先确认当前授权可用，恢复后校验快照与核心公式；专项恢复入口，不可盲目重跑。
- 本轮核对2026-11-04第6单及23:45第7单配送费：夜间库存步骤失败在配送费复查之前，待查队列并不代表费用已经入账。恢复与补同步仅以实际平台配送实付为准，不把待完成/查询失败当0元。
- `agent/runtime_paths.ps1` 新增 `Initialize-PetstarLarkUserEnvironment`，供计划任务及直接同步入口统一清除宿主路由标记并选择当前Windows用户授权目录，只影响子进程环境，不读/复制token。看门狗夜间恢复统一走仓库内计划任务入口，不再从旧Hermes专用脚本绕行；现有Windows任务无需重装。
- `work/sync_meituan_inventory.ps1` 与 `agent/sync_sku_master_by_identity.ps1` 的失败路径改为关闭/加入既有事务后抛出原异常，避免嵌套脚本 `exit` 吞掉根因；本地工作簿JSON显式按UTF-8读。失败仍回滚，禁止带半成品提交。
- `agent/sync_delivery_fees.ps1` 与 `agent/delivery_fee_policy.mjs` 每小时除当天全部自然单和持久待查队列外，额外发现前一天零/未知且未人工确认的费用；只加一个日期，不扫描全历史。费用后结算/前次任务中断均可续查。跨日补录仅刷新相关月份，在提交成功后发费用更正提醒；未知费用继续警告“非确认免费、利润可能偏高”，不能猜实付或改平台数据。
- `work/test_night_sync_recovery.ps1` 离线验证授权环境一致、嵌套错误透传、UTF-8读取及看门狗统一入口；配送Node/PowerShell回归覆盖跨日无队列补查及人工值保护，均不写真实经营表。
- 2026-11-05恢复记录：原事务09:28完成回滚；额外回读校验遇到字典/对象表示差异误报，已修正校验器并加回归，未据此改业务值。随后09:46补日结事务 `TX-20261105-093027-AACB38C0` 提交、11月一致性审计与公式校验通过，只初始化新订单成本并刷新2026-11，无全量重建。补齐23:45第7单 `3202374541264510287`：配送5.60元、利润约7.65元；店主确认第6单 `3402374350443155017` 自配送0元，独立事件09:47提交且未重算商品成本。当天订单利润16.81元，经营支出4.75元，净利润12.06元。
- 尚未解决的授权边界：09:48实际启动 `Petstar-Hourly-Sync`，后台仍返回用户凭据missing；同一Windows用户、相同授权目录，后台可见注册表条目2个，维护环境3个。统一入口本身不能证明真实后台可用，禁止宣称自动恢复；待店主同意后在真实桌面/任务环境授权，再实际运行任务验证。不复制token、不盲目重装任务或放宽权限。
- 更正与验证：只读重发不应把 `@($null)` 产生的空项当作跨日补录，已加空项/订单号/日期保护并验证（14项夜间专项、54项Node配送、25项PowerShell配送、61个入口可靠性及23项赠品回归通过）。09:53重发的11月4日日结无警告、净利润12.06元；两单费用来源均已回读确认。
- 店主已同意真实后台重新授权，临时任务 `Petstar-Lark-Reauthorize-20261105` 以当前用户交互登录、普通权限运行既有 `agent/authorize_lark_desktop.ps1`，请求/结果仅放本机临时目录，不纳入仓库或运行状态备份；已在桌面打开官方授权页、不展示授权详情。等待店主完成后必须再次实际启动每小时任务；验证后清理本轮临时任务及临时授权文件，原有同步任务保留。
- 授权帮助脚本增加 `-Resume` 续接原短期申请，原生stderr等待提示不得中断Windows PowerShell 5.1，失败仅保存安全错误分类（无凭据原文）。15项专项离线测试包含真实本地CMD stderr夹杂stdout JSON用例。首次授权任务提前退出，当前申请仍有效，等待店主在已打开页面完成后续接轮询；尚无后台授权成功结论，不自动重复创建授权申请。
- 上述“尚未恢复”状态已于2026-11-05 10:03更新：店主完成页面授权后，10:00真实Windows小时任务运行成功（LastTaskResult=0），最小读表验证成功，无变化轻量路径未重算成本或简账。跨日复查确认6号单人工自配送保护、7号单5.60元，待查0单。旧临时授权辅助结果文件仍可能显示先前失败，不能据此反复续接已过期申请；以实际后台授权读表及任务结果为验收依据。
- 调整前时段曾为09:00—22:00普通整点、23:00补录提醒并同步、00:00前一日日结，00:20/00:40仅日结失败补跑；此旧安排现已被上方“全天每小时同步”登记替代。配送实付未结时仍明确暂估、继续跨日追查，不承诺00点一定取得最终骑手实付。

### 2026-11-04 赠品具体规格定位与待确认恢复

- `agent/record_finance_event.ps1` 的 `AddOrderGift` 使用纯计划函数 `New-OrderGiftPlan`，接受普通KW/CSX或查询返回的 `sku_ref:行号:商品ID×数量`；安全引用逐次复核行号与商品ID，未知成本继续待核对，不以重复货号猜规格。赠品显示仍为当前KW，已定位的商品ID保存在原备注栏的系统标记中，不新增列、不改SKU单价或平台库存。
- `agent/rebuild_cost_snapshots.ps1` 优先按赠品备注中的稳定商品ID归属成本账，已存在该标记时不得退回可能改绑的KW；旧赠品无标记时保持原匹配规则。指定订单赠品确认后只更新该单及实际涉及SKU、所属月份；待同步赠品只暂存，预览不得触发成本或账表写入。
- `work/test_gift_sku_reference.ps1` 通过真实纯计划函数和内存表覆盖重复规格、失效引用、零/未知成本、数量、身份保留及幂等预览边界；禁止向飞书写测试数据。待确认恢复只允许新增该赠品预览及发送通知，保留既有干预待确认队列，不代替店主确认经营写入。
- 原失败消息已由既有重放机制于2026-11-04恢复为队列第2项 `0B8A6C14`，仅含 `AddOrderGift`，安全定位仍为主表第163行及商品ID51181657345，赠品成本1.40元；恢复当时第1项 `4022CA2B` 干预请求未执行、未覆盖，飞书通知已只读回查到。后续队列正常推进，最终核对时赠品已是当前待确认项；本轮未代替店主确认任何经营写入。普通“确认”仍按原先进先出规则处理。
- 23项赠品规格专项、批量补货安全定位、消息桥和运行可靠性离线回归通过；真实 `AddOrderGift -DryRun` 返回正确规格及1.40元，赠品记录与主表第163行前后只读比对一致，目标赠品仍为0条。未执行真实赠品记账，等待店主在飞书确认；当前订单类型以实时订单表为准，不沿用旧对话的自然单成本推算。
- 注意旧 `work/test_gift_order_formula.ps1` 会向正式订单汇总Z1:Z3写公式后清空，不属于离线测试；本轮尝试在CLI访问阶段失败、未发出写入，不得扩大权限盲目重跑。后续验收用零外部写入的纯函数回归及 `-DryRun` 前后只读比对，真实赠品入账仍需店主另行确认。

### 2026-11-04 已确认规格的批量补货安全定位

- `agent/record_finance_event.ps1` 的 `BatchSkuMaintenance` 补货项目接受查询返回的 `sku_ref:行号:商品ID`；逐次复核行号和商品ID，明确定位后不再按重复KW猜行，也不删除、解绑或修改其他同号档案。只读一次主表及补货记录，写入范围仍仅为目标SKU与新增补货行。
- 店主明确确认“快照已初始化的数量就是本次尚未登记的补货”时，可先用 `SetCostBookQuantity` 将对应补货前数量纠正为0，再在同一原子事务登记补货。不得把已初始化3盒再加成6盒；无此确认不得擅自纠正数量。普通KW不唯一、定位符失效或附带冲突CSX仍停止写入。
- `work/test_replenishment_sku_reference.ps1` 使用内存表及真实执行函数验证重复货号、指定新规格、旧行不变、数量纠正、成本计算及失效引用；不访问飞书、不修改经营数据。
- `work/record_confirmed_replenishments_20261104.ps1` 是本次6笔采购231元的专项恢复入口，默认只预演；`-Apply` 才在全局事务中将两款新盒装补货前数量纠正为0并登记6笔。依赖既有执行器、事务及飞书用户授权；校验旧罐装不变、最终各3盒、采购/付款总额231元和新增行格式，不改历史订单、简账或平台库存。成功提交后才标记原失败消息已恢复，防止后台重复处理。日期、SKU与事件键已固定，不可盲目重跑或用于其他批次。
- 已于2026-11-04恢复并提交本次经营事务：补货记录第55～60行，采购和付款均为231元；新盒装KW4171/4172各3盒、当前成本各约7.1167元，KW1033/1018各39.705元，KW1013/1034各27.37元。旧罐装第591/592行未变；原失败消息仅在提交成功后标为完成。KW重复状态仍保留，未擅自改号或清除冲突标记；8项内存回归、消息桥与运行可靠性测试通过。

### 2026-11-03 配送费配置编码修复与定向补同步

- `agent/sync_delivery_fees.ps1` 统一以UTF-8读取本地工作簿绑定、干预种子及配送待查队列，避免Windows PowerShell 5.1把无BOM中文JSON按系统代码页解析。读取失败仅报告文件用途，不回显配置原文或链接。
- `agent/run_sync_and_notify.ps1` 对配送费步骤分别提示本地数据、浏览器连接、登录和其他查询失败；不能把程序异常一律说成平台登录过期。每小时继续复查当日自然单，人工费用保护及未知金额不写0的规则不变。
- `work/delivery_fee_extract.mjs` 已以2026-11-03目标订单的详情及麦芽田页面 `data-detail-goods-template` 核实：“配送费用”直接取 `delivery.send_fee / 100`，非零 `discount_fee` 不应再次扣减或单凭其存在阻断入账。保险、报损、小费、多次派送、金额不一致仍待核对；回归包含387分配送费与38分优惠仍记3.87元，禁止错记3.49元。
- `work/test_delivery_fee_sync.ps1` 增加无BOM中文JSON、中文路径、损坏JSON安全报错及主流程提示分类回归；不得用虚构订单写飞书。
- `work/repair_delivery_fees_20261103.ps1` 为仅限本次日期的专项补同步入口，默认只读，`-Apply` 才加入全局事务写自然单配送费并刷新2026-11；写后比对全部既有订单商品成本未变、收入表费用/成本/利润一致。禁止盲目重跑，不用于其他日期；不运行全量重建、不改美团或配送平台数据。
- 首次专项核对因“汇总显示到分、收入保留小数”比较口径不同而触发安全回滚，已改为收入表同列前后原值比较，禁止降低金额精度来通过校验。12:00例行任务已于12:06自动完成真实配送费写入、仅刷新2026-11，两个平台查询和公式校验成功、成本重建未运行；不要再盲目重跑专项写入。离线回归47项Node费用用例、25项Windows PowerShell费用用例及运行可靠性检查通过。

### 2026-11-03 待办事项清单

- `work/ensure_todo_sheet.ps1` 在现有经营驾驶舱工作簿新增“待办事项”，四列为“日期 / 待办事项 / 状态 / 备注”。默认预览，`-Apply` 才创建和初始化空表；已有人工内容不覆盖、不删除。日期采用真实日期格式，状态为“未完成 / 已完成”。
- `agent/todo_store.ps1` 为待办专用读写及纯计划函数：新增只写一行，完成只改状态格；聊天消息ID幂等、写前快照、回读、公式校验均复用全局轻量事务，不调用商品成本或任何经营简账刷新。完成前必须唯一定位，查询返回绑定行号及内容哈希的 `todo_ref`，过期/同名/缺对象先澄清，禁止批量猜测完成。
- `agent/codex_response_schema.json`、`agent/codex_chat_bridge.py` 增加 `todos` 只读查询及 `AddTodo / CompleteTodo` 白名单。`agent/lark_event_bridge.ps1` 仅对纯待办动作直接执行并回复；与记账混合的消息要求拆分，本轮不扩大财务动作权限。聊天日期以原消息在Asia/Shanghai的日期为基准，未指定默认当天；表格手填日期留空持续提醒，状态留空按未完成。
- `agent/run_sync_and_notify.ps1` 的23:00补录文字提醒读取当前待办表，展示当日、历史遗留及无日期的未完成事项，未来事项及已完成事项不列出；读取失败明确提示，不能当作零待办，也不反向回滚已提交账目。现有任务时间不变，不增加新定时任务。
- `work/test_todo_store.ps1` 离线验证日期、状态、重复/过期定位、幂等、轻量写入及提醒；`work/test_lark_event_bridge.ps1` 验证纯待办直写、混合请求隔离及锁忙重试，使用独立测试事务锁，避免与真实整点同步相互阻塞；`work/test_codex_chat_bridge.py` 验证动作契约和手填状态不使用旧查询缓存。飞书上不创建虚构待办用于测试。上线需受控重载空闲消息监听器，恢复机制保留未处理消息；不修改密钥、登录态或Windows任务配置。

### 2026-11-03 可手动维护的搜索关键词库

- `work/ensure_keyword_library.ps1` 在现有经营驾驶舱工作簿创建 `关键词库`，四列为“启用 / 搜索关键词 / 商品品类 / 关联货号（选填）”；默认预览，`-Apply` 才建表和初始化。只给空表填种子，已有内容只校验、不覆盖。初始化从店主已提供的品类/品牌词出发，品牌词核对最新在售有库存商品后才启用；关联货号默认空，不猜商品。结构创建失败不删除工作表，可幂等重跑。
- `agent/keyword_library.ps1` 为纯读取解析、关键词匹配与轮换函数；`agent/send_daily_traffic_suggestion.ps1` 每次未发送的10点提醒从飞书读取已启用关键词，不回退写死词。按最近推荐历史轮换关键词，再选关联在售有库存商品；近期销量仍只作参考。关联货号填写后优先按货号限定；缺词、格式错误或无法关联商品时发明确运营检查提醒，不用无关商品冒充匹配。
- 日常只读 `关键词库`、库存快照、近期自然订单；只发正常商品展示/库存/价格检查建议，不安排补单人员或虚假交易，不修改订单、库存、SKU成本或账表。关键词只在成功发送后进入原有90天去重历史；预演不发消息、不占用下一次提醒。
- `work/test_keyword_library.ps1` 离线覆盖启停、手动新增、同词去重、品牌/品类与货号匹配、零销量商品及轮换；`work/test_runtime_reliability.ps1` 覆盖10点入口和词库依赖。现有 Windows 任务无需重装。
- 2026-11-03 已在原工作簿创建 `关键词库`（sheet `uIOSdE`）：13个初始词，启用12个；城市理享猫砂暂未核实到在售有库存商品而停用。已回读首/中/末行、样式、列宽、启停下拉和冻结表头，公式检查成功；10点入口 `-DryRun` 从飞书选词、匹配商品成功且未发消息、未写推荐历史；27项离线词库测试及160项可靠性契约通过。

### 2026-11-03 配送费逐小时复查

- `agent/sync_delivery_fees.ps1` 接入现有整点/日结事务：每次检查目标日全部有效美团自然单，不以费用已为0或曾查询过为跳过条件；跨天仍待发配送/待完成/冲突订单进入 `agent/state/delivery_fee_pending.json` 继续追踪，明确自配送或转为干预单后退出。
- `work/delivery_fee_extract.mjs` 通过现有 web-access CDP 在自建后台标签页只读查询麦芽田、青云聚信，按日期批量分页，并按完整美团订单号与店铺匹配；不输出顾客资料、Cookie、token，不发单、取消、付款、加价或修改配送。
- `agent/delivery_fee_policy.mjs` 负责候选选择与金额裁决：麦芽田优先、青云兜底；只接收完成配送的明确实付；两边正数冲突、重复记录、金额未知或多次完成派送保持待核对。人工值及实际履约覆盖优先，查不到/0不自动认作免费自配送。页面/授权失败只降级配送费步骤，不阻断订单同步。
- 写入只涉及指定订单汇总配送费及来源备注，由现有事务回滚保护、公式校验及回读验证，再刷新受影响月份；不重算商品成本、不改库存、不处理干预单费用或自动登记售后追加配送费。只读预演不会写账或持久化待查队列。
- `work/test_delivery_fee_sync.mjs` 和 `work/test_delivery_fee_sync.ps1` 为离线测试入口；不得用虚构订单向飞书写入。
- 已通过两个平台真实已登录页面的只读预演（6.78元、4.51元，与现有人工金额一致，零账表写入），以及 Windows PowerShell 5.1 / Node 离线回归。下次现有整点任务自动加载，无须修改 Windows 计划任务；尚未用真实费用变更执行写账验收。`sync_delivery_fees.ps1` 单独运行默认只读，`-Apply` 必须加入主同步已有事务，不能单独写完费用后遗漏月份刷新。

### 2026-11-03 补单日期误报修复

- `work/sync_meituan_orders.ps1` 的补单日期补救只比较本地自然日，保留已有时间，不再先用本次抓取的时分覆盖日期再当成异常修正；真正缺日期或日期跨天错误仍按订单汇总补救。
- `work/test_fast_sync_policy.ps1` 离线重放实际补单更新/补救代码，覆盖文本日期、带时分、Excel小数日期、缺失日期及真实跨日错误；同一天无日期修复、无费用/库存变化、无假变更月份。日期真修复重跑不重复提醒。
- 本轮只改同步判断及测试，不回写历史日期、不改手工配送费（包括已确认的4.51元）、不运行账表重建；下次计划任务自动加载新脚本。

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
| 每日引流款运营建议 | `agent/send_daily_traffic_suggestion.ps1`、`agent/keyword_library.ps1`、`agent/traffic_suggestion_rules.json` | 只读飞书关键词库、最新库存快照和最近 30 天自然订单；仅写本地去重历史并向当前授权用户发飞书私聊 | 每天 10:00 按飞书启用词和近期推荐历史轮换，再匹配在售有库存商品；明确关联货号时不改荐其他货号；近期销量只作参考。缺词/缺匹配发维护提醒，表格读取失败不偷偷回退固定词；只建议正常商品展示、库存和推广检查，不创建订单、不评价、不改库存 |
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

- `work/repair_order_costs_20261102.ps1`：仅修正订单 `3602371202674777015` 的30g羊奶粉（7元/罐）及订单 `2102365133521804156` 的16g×12猫条（17.64元/盒）成本快照；按完整订单号、商品名和规格唯一定位，默认只预览，`-Apply` 在全局事务内写入指定商品 L:N 并只刷新2026-10/11简账。读取订单、赠品、对应收入与公式；不填货号、不改SKU主表、库存、补货、退款或其他历史订单成本；有退款、未核成本或身份冲突时停止。专项修复，不可盲目重跑。

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
