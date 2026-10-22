from __future__ import annotations

import argparse
import base64
import collections
import concurrent.futures
import datetime as dt
import hashlib
import json
import os
import pathlib
import queue
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
from typing import Any


if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")


PROJECT_ROOT = pathlib.Path(__file__).resolve().parent.parent
STATE_ROOT = PROJECT_ROOT / "agent" / "state"
SESSION_STATE = STATE_ROOT / "codex_conversation_state.json"
EVENT_CACHE = STATE_ROOT / "codex_event_cache.json"
LOCK_PATH = STATE_ROOT / "codex_chat_bridge.lock"
RUNTIME_START_LOCK_PATH = STATE_ROOT / "codex_runtime_start.lock"
RUNTIME_STATE_PATH = STATE_ROOT / "codex_runtime_state.json"
PERFORMANCE_LOG_PATH = STATE_ROOT / "codex_performance.ndjson"
SCHEMA_PATH = PROJECT_ROOT / "agent" / "codex_response_schema.json"
QUERY_TOOL_PATH = PROJECT_ROOT / "agent" / "query_business_data.ps1"
WINDOWS_POWERSHELL = pathlib.Path(os.environ.get("SystemRoot", r"C:\Windows")) / "System32" / "WindowsPowerShell" / "v1.0" / "powershell.exe"
CODEX_JS = pathlib.Path(os.environ.get("APPDATA", "")) / "npm" / "node_modules" / "@openai" / "codex" / "bin" / "codex.js"
NODE_COMMAND = shutil.which("node.exe") or shutil.which("node") or "node"
CODEX_PREFIX = [NODE_COMMAND, str(CODEX_JS)]
MAX_SECONDS = 240
QUERY_SECONDS = 60
MAX_QUERY_ROUNDS = 8
MAX_TOTAL_QUERIES = 30
QUERY_CACHE_SECONDS = 30
RUNTIME_VERSION = "2026-10-09.1"
RUNTIME_PIPE = rf"\\.\pipe\petstar_codex_runtime_{hashlib.sha256(str(PROJECT_ROOT).encode('utf-8')).hexdigest()[:16]}"
RUNTIME_AUTHKEY = hashlib.sha256((str(PROJECT_ROOT) + "|petstar-codex-runtime-v1").encode("utf-8")).digest()
QUERY_DATASETS = {
    "overview",
    "orders",
    "order_items",
    "pending_costs",
    "interventions",
    "sku",
    "inventory",
    "replenishments",
    "income",
    "expenses",
    "profit",
    "status",
}
ANSI_RE = re.compile(r"\x1b\[[0-?]*[ -/]*[@-~]")
SECRET_PATTERNS = (
    re.compile(r"(?i)(bearer\s+)[a-z0-9._~-]+"),
    re.compile(r"(?i)(app_secret[\s:=]+)[^\s,;]+"),
    re.compile(r"(?i)(token[\s:=]+)[^\s,;]+"),
    re.compile(r"\bsk-[A-Za-z0-9_-]{8,}\b"),
)

_QUERY_CACHE: dict[str, tuple[float, dict[str, Any]]] = {}
_QUERY_CACHE_LOCK = threading.Lock()
_APP_SERVER_RUNTIME: "CodexAppRuntime | None" = None


def get_auth_fingerprint() -> str:
    """Return a non-secret generation marker for the active Codex credentials."""
    codex_home = pathlib.Path(os.environ.get("CODEX_HOME") or (pathlib.Path.home() / ".codex"))
    auth_path = codex_home / "auth.json"
    try:
        payload = auth_path.read_bytes()
        stat = auth_path.stat()
        return hashlib.sha256(
            payload + f"|{stat.st_mtime_ns}|{stat.st_size}".encode("ascii")
        ).hexdigest()
    except OSError:
        return "missing"


def is_authentication_error(error: BaseException) -> bool:
    text = str(error).lower()
    return any(
        marker in text
        for marker in (
            "access token",
            "refresh token",
            "logged out",
            "signed in with another account",
            "authentication",
            "unauthorized",
            "not logged in",
        )
    )

ACTION_GUIDE = """
可返回的动作（只填店主明确提供或能无歧义推导的参数）：
- AddIntervention: date, order_no, person, note
- RemoveIntervention: order_no, note。仅当店主明确说某订单不是干预单、登记错了或要求撤销干预时使用
- RestoreInventory: order_no, note
- ConfirmSelfDelivery: order_no, note
- SetExternalDeliveryFee: order_no, amount, note
- SetOrderDeliveryFee: order_no, amount, note
- SetOrderActuals: order_no, amount（最终实际收入）, items（最终实际交付的KW/CSX×数量）, delivery_fee, note。用于线上订单私下退款、漏送退款、到店换货、实际成交商品变化或最终改价；漏送/退回的商品不得写入 items；该动作不修改平台库存
- OfflineSale: date, items(KW/CSX×数量), amount, delivery_fee, receipt_status, platform, note。receipt_status 只填“已到账”或“未到账”；platform 只填“美团”“淘宝闪购”“京东到家”“线下”（用户说“京东外卖”时填“京东到家”）
- AddOrderGift: date, order_no或daily_order_no, items, note
- Replenishment: date, goods_code, transition_code（消息同时给出CSX时填写）, quantity, purchase_amount, payment_amount, payment_date, production_date, shelf_life_value, shelf_life_unit, note。该动作只新增补货记录并更新SKU当前成本，不重算历史订单、收入、利润或经营简账
- CashExpense / OperatingExpense / OwnerWithdrawal / OwnerLabor / MarketRent / BeginningCash / ActualReceipt / AssetAmortization: date, amount, note。ActualReceipt 只表示平台资金实际打入收款账户的月度现金到账，不得用于修正某一订单的收入、退款或利润，也不得携带 order_no
- SetSkuExpiry: goods_code, expiry_date；明确无到期日时 clear_expiry=true
- SetSkuBatchDate: goods_code, production_date, shelf_life_value, shelf_life_unit
- SetSkuCost: goods_code, amount（每个销售单位的新成本，允许为0）, order_no（仅当店主明确指定某个历史订单并明确要求同步修正该单成本/利润时填写）, note；goods_code 优先 KW/CSX，也可使用 pending_costs 返回且能唯一定位主表的 action_code。普通成本修改不得填写 order_no，不得额外生成 RefreshLedger
- ChangeSkuCode: goods_code(旧号), new_goods_code；若明确旧商品已下架/移走可 retire_existing_new_code=true
- SetRefundDisposition: order_no, goods_code, refund_disposition。已退回且恢复库存填 returned_restocked；退款但未退货填 refund_no_return；退回但损坏不能再售填 returned_damaged
- RefreshLedger: date
- SendDailySummary: date, mode固定night
- SyncPlatformData: date
""".strip()

BASE_GUIDANCE = """
你是“宠速星经营管家”的经营决策层，服务一家宠物用品即时零售店。飞书只是店主与你沟通的入口，你要像一个真正了解本店数据的经营助理一样理解、查询、回答和安排动作。
店主可以使用任意自然语言、口语、省略句和一次包含多个事项，不要求固定命令格式。
你负责理解、读取经营数据、分析回答和生成白名单动作；绝不直接修改飞书、美团、库存、订单或本地经营数据，也不要运行任何写入脚本。真正写入由外层桥接器在校验和确认后执行。

当问题涉及本店已经保存的订单、干预单、收入、支出、库存、SKU、补货、成本或同步状态时，必须先在 queries 中声明需要的数据集和过滤条件，不能凭对话猜测，也不能说“看不到表格”或要求店主重新提供系统已有的数据。可用 dataset：overview、orders、order_items、pending_costs、interventions、sku、inventory、replenishments、income、expenses、profit、status。日期使用 yyyy-MM-dd；尽量用最小 limit。外层只读工具会读取数据并再次交给你，你可以根据结果继续调用另一个只读工具，证据充分后才回答。你自己不得启动 PowerShell、lark-cli、经营脚本或网页自动化。查询失败时要说明是哪个数据域读取失败，不得编造结果。
精确订单商品明细使用 order_items；成本未知、历史订单成本修正或用户说“这一单未知成本的商品”时使用 pending_costs，并优先按完整订单号过滤；指定时间盈利使用 profit。pending_costs 会给出 current_kw、transition_code、action_code、标题、数量、主表行号和成本状态：能唯一得到 action_code 时，用户明确给出单位成本后可生成 SetSkuCost；没有 action_code 时才追问当前 KW/CSX。只有店主明确要求更新该历史订单的成本或利润时，SetSkuCost 才携带查询得到的完整 order_no；只说“修改某SKU成本”时只改主表，绝不能联动历史订单或简账。只要待核成本查询有结果，回复或待确认预览必须逐项展示：当前 KW（没有则显示过渡货号）、商品标题、该订单数量、SKU成本主表行号、当前成本/状态、新单位成本，以及是否联动该订单；不得只列货号和行号。
纯查询直接给结论且 actions 为空；查询和修改混在同一条消息时，先基于查询结果回答，再清楚预览准备修改的内容。列表尽量使用短项目符号，金额与统计口径要写清楚。

当前消息是唯一的新行动来源。历史对话仅帮助理解“这单、昨天、刚才那个”等指代，禁止重放历史动作，禁止把旧待办当成当前请求。
订单号必须逐字保留，不能改写、补全或猜测。KW 是可调整和复用的当前货号，不是永久商品身份；不得创建历史别名。
店主纠正“旧订单不是干预单，真正的干预单是另一个订单”时，必须为旧订单生成 RemoveIntervention，并为正确订单生成 AddIntervention；配送方式、配送费、补单人和库存恢复状态如果是在同一纠错上下文中明确给出的，也一并迁移到正确订单。不得只解释执行层不支持，也不得把两笔都保留为干预。
自己送、自配送、商家自配、没有配送费，都表示配送费为0元。线下单默认自提配送费0并当天到账，除非店主另有说明。
线上订单发生私下退款、漏送退款、到店换货、实际售出商品与线上订单不同，或店主明确给出该订单最终实收时，使用 SetOrderActuals，绝不能使用 ActualReceipt。amount 填最终实际收入；items 只填最终实际交付商品，漏送、退回或没有交付的商品必须排除。用户没有明确谈配送方式或配送费时 delivery_fee 必须填 null，由执行器在确认时沿用订单当前值；只有明确到店自提、自配送或给出新配送费时才填写。店主说已手工恢复美团库存时只写进 note，绝不能再生成库存修改动作。若原订单商品和数量全部保留、仅私下加购一个新商品并明确额外付款，则 amount=原商家预计收入+额外付款，items=原商品+新增商品，不要追问金额分摊；多个新增/替换商品且无法确定售价分摊时才追问，不能猜。
干预单不计算购买商品进货成本；自然单和线下单计算货物成本。赠品成本按登记时成本快照计入。
生产日期加保质期可推算到期日；缺少保质期不得猜。今天、昨天、前天以 Asia/Shanghai 和消息创建时间为准。
只要当前消息明确表达“补货/进货/采购”并提供数量，就优先生成 Replenishment；即使同时出现生产日期、保质期，也不能误判成仅修改已有批次。只有没有新增数量和采购事实、明确是在修正已有库存日期时，才使用 SetSkuBatchDate。
店主明确说试吃装、赠送、免费或成本为0时，Replenishment 的 purchase_amount 可以是0，并在 note 保留零成本原因。新KW尚未存在于SKU成本主表时也照常生成 Replenishment，由执行层建立待平台同步档案，不能改成 SetSkuBatchDate 或直接报无法处理。
“补货”既可能是已有商品再次采购，也可能是第一次上新的入库，不能要求店主换一种固定说法。消息同时给出 KW 和 CSX 时，goods_code 填 KW、transition_code 填 CSX；只有 CSX 时 goods_code 填 CSX。执行层会优先用任一编号唯一匹配 SKU 主表，并在无冲突时把新 KW 绑定到该 CSX 所在档案。
补货只更新补货记录和该SKU当前成本；即使一次有多条补货，也不得生成 RefreshLedger，不得声称会修改库存批次、现金流、历史订单、收入、利润或经营简账。
店主要求修改、补充或纠正某个 KW/CSX 的成本价时，使用 SetSkuCost。amount 是该商品每个“销售单位”的成本，不是整箱金额；若无法判断整箱与单件换算，只追问每箱包含多少销售单位，不能猜。默认只修改SKU成本主表；只有店主明确指定历史订单并要求联动该单时才填写 order_no，且只允许更新该订单，不得扩展为全量历史重算。

回复必须简洁、自然，像一个懂店铺的同事，不输出推理、JSON、脚本、字段清单、内部路径、技术校验或英文内部状态。通常不超过4行、220个汉字。无论成功理解、需要追问还是失败，都必须给 reply。
""".strip()


def now_iso() -> str:
    return dt.datetime.now().astimezone().isoformat()


def load_json(path: pathlib.Path, default: Any) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (FileNotFoundError, json.JSONDecodeError, OSError):
        return default


def save_json(path: pathlib.Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = path.with_suffix(path.suffix + ".tmp")
    temp_path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    os.replace(temp_path, path)


def append_performance(event_key: str, mode: str, ok: bool, timings: dict[str, Any]) -> None:
    """Append non-sensitive stage timings; never let telemetry break message handling."""
    try:
        PERFORMANCE_LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
        record = {
            "recorded_at": now_iso(),
            "event_hash": hashlib.sha256(event_key.encode("utf-8")).hexdigest()[:16],
            "mode": mode,
            "ok": ok,
            **timings,
        }
        with PERFORMANCE_LOG_PATH.open("a", encoding="utf-8") as handle:
            handle.write(json.dumps(record, ensure_ascii=False, separators=(",", ":")) + "\n")
    except OSError:
        pass


def clear_query_cache() -> int:
    with _QUERY_CACHE_LOCK:
        count = len(_QUERY_CACHE)
        _QUERY_CACHE.clear()
    return count


def output(payload: dict[str, Any]) -> int:
    print(json.dumps(payload, ensure_ascii=False))
    return 0 if payload.get("ok") else 1


def decode_prompt(value: str) -> str:
    try:
        return base64.b64decode(value, validate=True).decode("utf-8").strip()
    except (ValueError, UnicodeDecodeError) as exc:
        raise ValueError("消息内容编码无效") from exc


def redact_error(value: str) -> str:
    text = ANSI_RE.sub("", value or "").strip()
    for pattern in SECRET_PATTERNS:
        text = pattern.sub(
            lambda match: (match.group(1) if match.lastindex else "") + "[已隐藏]",
            text,
        )
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    return "；".join(lines[-5:])[:1000]


def event_cache_key(event_key: str, mode: str) -> str:
    return hashlib.sha256(f"{mode}:{event_key}".encode("utf-8")).hexdigest()


def get_cached_response(event_key: str, mode: str) -> dict[str, Any] | None:
    cache = load_json(EVENT_CACHE, {})
    entry = cache.get(event_cache_key(event_key, mode), {})
    if entry.get("status") == "completed" and isinstance(entry.get("result"), dict):
        return entry["result"]
    return None


def cache_response(event_key: str, mode: str, result: dict[str, Any]) -> None:
    cache = load_json(EVENT_CACHE, {})
    cache[event_cache_key(event_key, mode)] = {
        "status": "completed",
        "result": result,
        "processed_at": now_iso(),
    }
    ordered = sorted(
        cache.items(),
        key=lambda item: str(item[1].get("processed_at") or ""),
        reverse=True,
    )[:300]
    save_json(EVENT_CACHE, dict(ordered))


class ProcessLock:
    def __init__(self, path: pathlib.Path, timeout_seconds: int = 300) -> None:
        self.path = path
        self.timeout_seconds = timeout_seconds
        self.handle: Any = None

    def __enter__(self) -> "ProcessLock":
        import msvcrt

        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.handle = self.path.open("a+b")
        if self.handle.tell() == 0:
            self.handle.write(b"0")
            self.handle.flush()
        deadline = time.monotonic() + self.timeout_seconds
        while True:
            try:
                self.handle.seek(0)
                msvcrt.locking(self.handle.fileno(), msvcrt.LK_NBLCK, 1)
                return self
            except OSError:
                if time.monotonic() >= deadline:
                    self.handle.close()
                    self.handle = None
                    raise TimeoutError("本机 Codex 正在处理上一条消息，等待超时")
                time.sleep(0.25)

    def __exit__(self, exc_type: Any, exc: Any, traceback: Any) -> None:
        if not self.handle:
            return
        import msvcrt

        try:
            self.handle.seek(0)
            msvcrt.locking(self.handle.fileno(), msvcrt.LK_UNLCK, 1)
        finally:
            self.handle.close()
            self.handle = None


def build_prompt(
    prompt: str,
    recent_turns: list[dict[str, str]],
    mode: str,
    event_key: str,
) -> str:
    context_lines: list[str] = []
    for turn in recent_turns[-8:]:
        role = "店主" if turn.get("role") == "user" else "管家"
        content = str(turn.get("content") or "").strip()
        if content:
            context_lines.append(f"{role}：{content}")
    context = "\n".join(context_lines) if context_lines else "（无）"
    if mode == "preview":
        phase = """
当前是待确认预览。若请求会新增、修改或覆盖经营数据：信息齐全时返回 actions；只有 actions 非空时 reply 第一行才写“📝 待确认”，并逐项说明要改什么，不能只写“待确认”。actions 为空时禁止使用“待确认”“准备修改”“确认后执行”等措辞。涉及 SKU 成本时，每个 SKU 都要写明当前 KW（没有则写过渡货号）、商品标题、订单数量、SKU成本主表行号、旧成本/状态、新单位成本，并明确“只改SKU成本主表”或“按店主明确要求联动订单号……”。不得自行扩大联动范围。缺关键事实时 actions 必须为空、needs_information=true，只追问一个最关键问题，不要声称已写入。
""".strip()
    elif mode == "execute":
        phase = """
店主已确认刚才的预览。请只为当前原始请求重新生成与预览一致的 actions，不扩大范围。信息不足则不要生成动作。reply 只概括将执行的事项，不得提前声称成功。
""".strip()
    else:
        phase = """
当前是普通对话。查询、解释可直接回答且 actions 为空。只有“刷新简账、补发日结、同步平台数据”可在普通对话返回对应维护动作。若内容要求修改经营数据，返回动作及待确认说明，不得声称已执行。
""".strip()
    return f"""{BASE_GUIDANCE}

{ACTION_GUIDE}

{phase}

消息事件标识：{event_key}
当前阶段：{mode}
当前时间：{now_iso()}

近期对话（只用于理解指代，禁止当作新动作）：
{context}

店主当前原话（唯一的新请求）：
{prompt}
"""


def normalize_result(value: Any) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise ValueError("Codex 未返回对象")
    reply = str(value.get("reply") or "").strip()
    if not reply:
        raise ValueError("Codex 没有生成回复")
    actions = value.get("actions")
    if not isinstance(actions, list):
        actions = []
    queries = value.get("queries")
    if not isinstance(queries, list):
        queries = []
    if len(reply) > 1200:
        reply = reply[:1199].rstrip() + "…"
    return {
        "reply": reply,
        "needs_information": bool(value.get("needs_information")),
        "queries": queries[:6],
        "actions": actions[:20],
    }


def _run_business_query(index: int, query: dict[str, Any], use_cache: bool) -> dict[str, Any]:
    if not isinstance(query, dict):
        return {"query_index": index, "ok": False, "error": "查询参数不是对象"}
    dataset = str(query.get("dataset") or "").strip().lower()
    if dataset not in QUERY_DATASETS:
        return {"query_index": index, "ok": False, "dataset": dataset, "error": "数据集不在白名单"}
    fingerprint = query_fingerprint(query)
    if use_cache:
        with _QUERY_CACHE_LOCK:
            cached = _QUERY_CACHE.get(fingerprint)
            if cached and time.monotonic() - cached[0] <= QUERY_CACHE_SECONDS:
                return {"query_index": index, **cached[1], "cache_hit": True}

    command = [
        str(WINDOWS_POWERSHELL),
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(QUERY_TOOL_PATH),
        "-Dataset",
        dataset,
    ]
    argument_map = (
        ("start_date", "-StartDate"),
        ("end_date", "-EndDate"),
        ("order_no", "-OrderNo"),
        ("code", "-Code"),
        ("search", "-Search"),
    )
    for field, flag in argument_map:
        value = query.get(field)
        if value is None or str(value).strip() == "":
            continue
        text_value = str(value).strip()
        if field in {"start_date", "end_date"} and not re.fullmatch(r"\d{4}-\d{2}-\d{2}", text_value):
            return {"query_index": index, "ok": False, "dataset": dataset, "error": f"{field} 日期格式无效"}
        if len(text_value) > 200:
            return {"query_index": index, "ok": False, "dataset": dataset, "error": f"{field} 过长"}
        command.extend([flag, text_value])
    raw_limit = query.get("limit")
    try:
        limit = 100 if raw_limit is None else max(1, min(int(raw_limit), 100))
    except (TypeError, ValueError):
        return {"query_index": index, "ok": False, "dataset": dataset, "error": "limit 无效"}
    command.extend(["-Limit", str(limit)])
    try:
        completed = subprocess.run(
            command,
            cwd=str(PROJECT_ROOT),
            text=True,
            encoding="utf-8-sig",
            errors="replace",
            capture_output=True,
            timeout=QUERY_SECONDS,
            check=False,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        payload = json.loads((completed.stdout or "").strip())
        if completed.returncode != 0 or not payload.get("ok"):
            raise RuntimeError(payload.get("error") or completed.stderr or "查询命令失败")
        cache_payload = dict(payload)
        if use_cache:
            with _QUERY_CACHE_LOCK:
                _QUERY_CACHE[fingerprint] = (time.monotonic(), cache_payload)
        return {"query_index": index, **cache_payload, "cache_hit": False}
    except subprocess.TimeoutExpired:
        return {"query_index": index, "ok": False, "dataset": dataset, "error": "只读查询超时"}
    except Exception as exc:
        return {"query_index": index, "ok": False, "dataset": dataset, "error": redact_error(str(exc))}


def run_business_queries(queries: list[dict[str, Any]], use_cache: bool = True) -> list[dict[str, Any]]:
    """Run independent read-only Feishu queries concurrently while preserving request order."""
    selected = list(queries[:6])
    if not selected:
        return []
    results: list[dict[str, Any]] = []
    with concurrent.futures.ThreadPoolExecutor(max_workers=min(3, len(selected))) as executor:
        futures = [executor.submit(_run_business_query, index, query, use_cache) for index, query in enumerate(selected, start=1)]
        for future in futures:
            results.append(future.result())
    return results


def query_fingerprint(query: dict[str, Any]) -> str:
    normalized = {
        key: query.get(key)
        for key in ("dataset", "start_date", "end_date", "order_no", "code", "search", "limit")
    }
    return json.dumps(normalized, ensure_ascii=False, sort_keys=True, separators=(",", ":"))


def parse_jsonl(stdout: str) -> tuple[str, str]:
    thread_id = ""
    final_text = ""
    for raw_line in stdout.splitlines():
        line = raw_line.strip()
        if not line.startswith("{"):
            continue
        try:
            event = json.loads(line)
        except json.JSONDecodeError:
            continue
        if event.get("type") == "thread.started" and event.get("thread_id"):
            thread_id = str(event["thread_id"])
        item = event.get("item") or {}
        if event.get("type") == "item.completed" and item.get("type") == "agent_message":
            final_text = str(item.get("text") or "")
    return thread_id, final_text


class CodexAppRuntime:
    """One persistent Codex app-server process for all Feishu messages."""

    def __init__(self) -> None:
        self.process: subprocess.Popen[str] | None = None
        self.messages: queue.Queue[dict[str, Any]] = queue.Queue()
        self.stderr_tail: collections.deque[str] = collections.deque(maxlen=40)
        self.pending: collections.deque[dict[str, Any]] = collections.deque()
        self.next_id = 0
        self.lock = threading.Lock()
        self.schema = load_json(SCHEMA_PATH, {})
        self.auth_fingerprint = ""
        self._start()

    def _start(self) -> None:
        self.close()
        self.messages = queue.Queue()
        self.pending.clear()
        self.stderr_tail.clear()
        self.process = subprocess.Popen(
            [*CODEX_PREFIX, "app-server", "--stdio", "-c", "mcp_servers={}"],
            cwd=str(PROJECT_ROOT),
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            errors="replace",
            bufsize=1,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        threading.Thread(target=self._read_stdout, daemon=True).start()
        threading.Thread(target=self._read_stderr, daemon=True).start()
        response = self._request(
            "initialize",
            {"clientInfo": {"name": "petstar-runtime", "title": "Petstar Codex Runtime", "version": "1.0.0"}},
            30,
        )
        if "result" not in response:
            raise RuntimeError("Codex app-server 初始化失败")
        self._notify("initialized", {})
        self.auth_fingerprint = get_auth_fingerprint()

    def ensure_current_auth(self) -> bool:
        current = get_auth_fingerprint()
        if current == self.auth_fingerprint:
            return False
        self._start()
        return True

    def _read_stdout(self) -> None:
        assert self.process is not None and self.process.stdout is not None
        for raw in self.process.stdout:
            try:
                self.messages.put(json.loads(raw))
            except json.JSONDecodeError:
                continue

    def _read_stderr(self) -> None:
        assert self.process is not None and self.process.stderr is not None
        for raw in self.process.stderr:
            text = raw.strip()
            if text:
                self.stderr_tail.append(text)

    def _send(self, payload: dict[str, Any]) -> None:
        if self.process is None or self.process.poll() is not None or self.process.stdin is None:
            raise RuntimeError("Codex app-server 未运行")
        self.process.stdin.write(json.dumps(payload, ensure_ascii=False, separators=(",", ":")) + "\n")
        self.process.stdin.flush()

    def _notify(self, method: str, params: dict[str, Any]) -> None:
        self._send({"jsonrpc": "2.0", "method": method, "params": params})

    def _request(self, method: str, params: dict[str, Any], timeout: int) -> dict[str, Any]:
        self.next_id += 1
        request_id = self.next_id
        self._send({"jsonrpc": "2.0", "id": request_id, "method": method, "params": params})
        deadline = time.monotonic() + timeout
        while True:
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                raise subprocess.TimeoutExpired(method, timeout)
            try:
                message = self.messages.get(timeout=remaining)
            except queue.Empty as exc:
                raise subprocess.TimeoutExpired(method, timeout) from exc
            if message.get("id") == request_id:
                if message.get("error"):
                    raise RuntimeError(redact_error(json.dumps(message["error"], ensure_ascii=False)))
                return message
            self.pending.append(message)

    def _next_message(self, deadline: float) -> dict[str, Any]:
        if self.pending:
            return self.pending.popleft()
        remaining = deadline - time.monotonic()
        if remaining <= 0:
            raise subprocess.TimeoutExpired("turn/start", MAX_SECONDS)
        try:
            return self.messages.get(timeout=remaining)
        except queue.Empty as exc:
            raise subprocess.TimeoutExpired("turn/start", MAX_SECONDS) from exc

    def invoke(self, guided_prompt: str, thread_id: str = "") -> tuple[dict[str, Any], str]:
        with self.lock:
            self.ensure_current_auth()
            try:
                return self._invoke_locked(guided_prompt, thread_id)
            except RuntimeError as exc:
                if not is_authentication_error(exc):
                    raise
                # Account switching rotates the desktop/CLI OAuth grant. A
                # running app-server cannot adopt that grant in place, so start
                # it once with the latest protected credential file and retry.
                self._start()
                return self._invoke_locked(guided_prompt, thread_id)

    def _invoke_locked(self, guided_prompt: str, thread_id: str = "") -> tuple[dict[str, Any], str]:
        if self.process is None or self.process.poll() is not None:
            self._start()
        if not thread_id:
            started = self._request(
                "thread/start",
                {
                    "cwd": str(PROJECT_ROOT),
                    "ephemeral": True,
                    "approvalPolicy": "never",
                    "sandbox": "read-only",
                },
                45,
            )
            thread_id = str(started.get("result", {}).get("thread", {}).get("id") or "")
            if not thread_id:
                raise RuntimeError("Codex app-server 没有返回会话标识")
        response = self._request(
            "turn/start",
            {
                "threadId": thread_id,
                "input": [{"type": "text", "text": guided_prompt}],
                "cwd": str(PROJECT_ROOT),
                "approvalPolicy": "never",
                "sandboxPolicy": {"type": "readOnly"},
                "outputSchema": self.schema,
            },
            45,
        )
        turn_id = str(response.get("result", {}).get("turn", {}).get("id") or "")
        final_text = ""
        deadline = time.monotonic() + MAX_SECONDS
        while True:
            message = self._next_message(deadline)
            params = message.get("params") or {}
            if str(params.get("threadId") or "") != thread_id:
                continue
            if message.get("method") == "item/completed":
                item = params.get("item") or {}
                if item.get("type") == "agentMessage" and (not turn_id or str(params.get("turnId") or "") == turn_id):
                    final_text = str(item.get("text") or "")
            if message.get("method") == "turn/completed" and (not turn_id or str(params.get("turn", {}).get("id") or "") == turn_id):
                turn = params.get("turn") or {}
                if turn.get("status") != "completed":
                    raise RuntimeError(redact_error(json.dumps(turn.get("error") or turn, ensure_ascii=False)))
                break
        if not final_text:
            raise RuntimeError("Codex app-server 没有生成最终回复")
        try:
            parsed = json.loads(final_text)
        except json.JSONDecodeError as exc:
            raise RuntimeError("Codex 返回结果不是有效 JSON") from exc
        return normalize_result(parsed), thread_id

    def close(self) -> None:
        process = self.process
        self.process = None
        if process is None:
            return
        try:
            process.terminate()
            process.wait(timeout=5)
        except Exception:
            try:
                process.kill()
            except Exception:
                pass


def invoke_codex(guided_prompt: str, session_id: str = "") -> tuple[dict[str, Any], str]:
    if _APP_SERVER_RUNTIME is not None:
        return _APP_SERVER_RUNTIME.invoke(guided_prompt, session_id)
    STATE_ROOT.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(
        mode="w", encoding="utf-8", suffix=".json", dir=STATE_ROOT, delete=False
    ) as output_file:
        output_path = pathlib.Path(output_file.name)
    try:
        if session_id:
            command = [
                *CODEX_PREFIX,
                "exec",
                "resume",
                "--json",
                "--skip-git-repo-check",
                "--ignore-user-config",
                "--output-schema",
                str(SCHEMA_PATH),
                "--output-last-message",
                str(output_path),
                session_id,
                "-",
            ]
        else:
            command = [
                *CODEX_PREFIX,
                "exec",
                "--json",
                "--sandbox",
                "read-only",
                "--skip-git-repo-check",
                "--ignore-user-config",
                "--output-schema",
                str(SCHEMA_PATH),
                "--output-last-message",
                str(output_path),
                "-",
            ]
        creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        completed = subprocess.run(
            command,
            cwd=str(PROJECT_ROOT),
            input=guided_prompt,
            text=True,
            encoding="utf-8",
            errors="replace",
            capture_output=True,
            timeout=MAX_SECONDS,
            check=False,
            creationflags=creationflags,
        )
        next_session_id, event_text = parse_jsonl(completed.stdout or "")
        final_text = output_path.read_text(encoding="utf-8").strip() if output_path.exists() else ""
        if not final_text:
            final_text = event_text.strip()
        if completed.returncode != 0 or not final_text:
            detail = redact_error(completed.stderr or completed.stdout)
            raise RuntimeError(detail or f"Codex 返回退出码 {completed.returncode}")
        try:
            parsed = json.loads(final_text)
        except json.JSONDecodeError as exc:
            raise RuntimeError("Codex 返回结果不是有效 JSON") from exc
        return normalize_result(parsed), next_session_id or session_id
    finally:
        try:
            output_path.unlink(missing_ok=True)
        except OSError:
            pass


def check_runtime() -> dict[str, Any]:
    try:
        version = subprocess.run(
            [*CODEX_PREFIX, "--version"],
            cwd=str(PROJECT_ROOT),
            text=True,
            encoding="utf-8",
            errors="replace",
            capture_output=True,
            timeout=20,
            check=False,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        login = subprocess.run(
            [*CODEX_PREFIX, "login", "status"],
            cwd=str(PROJECT_ROOT),
            text=True,
            encoding="utf-8",
            errors="replace",
            capture_output=True,
            timeout=20,
            check=False,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        state = load_json(SESSION_STATE, {})
        return {
            "ok": version.returncode == 0 and login.returncode == 0 and SCHEMA_PATH.is_file() and CODEX_JS.is_file() and QUERY_TOOL_PATH.is_file(),
            "runtime": (version.stdout or version.stderr).strip(),
            "authenticated": login.returncode == 0,
            "schema_exists": SCHEMA_PATH.is_file(),
            "codex_entry_exists": CODEX_JS.is_file(),
            "query_tool_exists": QUERY_TOOL_PATH.is_file(),
            "session_ready": isinstance(state.get("recent_turns"), list),
            "conversation_mode": "isolated_message_session_with_explicit_recent_context",
        }
    except Exception as exc:
        return {"ok": False, "error": redact_error(str(exc))}


def process_message(prompt: str, event_key: str, mode: str = "normal") -> dict[str, Any]:
    started_at = time.perf_counter()
    cached = get_cached_response(event_key, mode)
    if cached:
        timings = {"total_seconds": round(time.perf_counter() - started_at, 3), "event_cache_hit": True}
        append_performance(event_key, mode, True, timings)
        return {"ok": True, "cached": True, **cached, "_timings": timings}

    lock_started_at = time.perf_counter()
    with ProcessLock(LOCK_PATH):
        lock_wait_seconds = time.perf_counter() - lock_started_at
        cached = get_cached_response(event_key, mode)
        if cached:
            timings = {
                "lock_wait_seconds": round(lock_wait_seconds, 3),
                "total_seconds": round(time.perf_counter() - started_at, 3),
                "event_cache_hit": True,
            }
            append_performance(event_key, mode, True, timings)
            return {"ok": True, "cached": True, **cached, "_timings": timings}
        session_state = load_json(SESSION_STATE, {})
        recent_turns = session_state.get("recent_turns") or []
        if not isinstance(recent_turns, list):
            recent_turns = []
        # Different inbound messages must not share Codex's hidden thread state.
        # We preserve conversational continuity explicitly through recent_turns,
        # while the temporary session is reused only inside this message's tool loop.
        session_id = ""
        guided_prompt = build_prompt(prompt, recent_turns, mode, event_key)
        model_seconds = 0.0
        query_seconds = 0.0
        query_cache_hits = 0

        def timed_invoke(text: str, current_session_id: str = "") -> tuple[dict[str, Any], str]:
            nonlocal model_seconds
            model_started_at = time.perf_counter()
            try:
                return invoke_codex(text, current_session_id)
            finally:
                model_seconds += time.perf_counter() - model_started_at

        try:
            result, next_session_id = timed_invoke(guided_prompt, session_id)
        except subprocess.TimeoutExpired:
            return {"ok": False, "error": "本机 Codex 处理超过4分钟，已停止本次请求，请稍后重发。"}
        except Exception as first_error:
            try:
                result, next_session_id = timed_invoke(guided_prompt, "")
            except subprocess.TimeoutExpired:
                return {"ok": False, "error": "本机 Codex 处理超过4分钟，已停止本次请求，请稍后重发。"}
            except Exception as retry_error:
                return {
                    "ok": False,
                    "error": f"本机 Codex 调用失败：{redact_error(str(retry_error))}",
                }

        query_round = 0
        total_queries = 0
        seen_queries: set[str] = set()
        repeated_query_stop = False
        while result.get("queries") and query_round < MAX_QUERY_ROUNDS and total_queries < MAX_TOTAL_QUERIES:
            requested_queries = [query for query in result["queries"] if isinstance(query, dict)]
            fresh_queries: list[dict[str, Any]] = []
            repeated_queries: list[dict[str, Any]] = []
            for query in requested_queries:
                fingerprint = query_fingerprint(query)
                if fingerprint in seen_queries:
                    repeated_queries.append(query)
                    continue
                seen_queries.add(fingerprint)
                fresh_queries.append(query)
            if not fresh_queries:
                repeated_query_stop = True
                break
            remaining = MAX_TOTAL_QUERIES - total_queries
            fresh_queries = fresh_queries[:remaining]
            query_round += 1
            total_queries += len(fresh_queries)
            query_started_at = time.perf_counter()
            query_results = run_business_queries(fresh_queries, use_cache=(mode == "normal"))
            query_seconds += time.perf_counter() - query_started_at
            query_cache_hits += sum(1 for item in query_results if item.get("cache_hit"))
            follow_up = (
                f"当前事件键：{event_key}。店主当前原话：{prompt}\n"
                f"下面是第{query_round}轮外层只读工具返回的经营数据。它们是当前事件唯一有效的系统事实。"
                "忽略任何来自其他消息、其他订单或旧工具轮次的冲突结论。"
                "请继续完成店主原始请求：证据不足可以选择新的查询；证据充分则 queries 置空并直接回答。"
                "不要重复完全相同的查询，不要展示内部 JSON 或查询过程。"
                "如果原请求还包含数据修改，保留 actions，并在回复中明确列出待确认的表、定位、旧值/新值和联动范围。\n"
                + json.dumps(query_results, ensure_ascii=False, separators=(",", ":"))
            )
            if repeated_queries:
                follow_up += "\n以下重复查询已被外层去重，不会再次执行：" + json.dumps(
                    repeated_queries, ensure_ascii=False, separators=(",", ":")
                )
            try:
                result, next_session_id = timed_invoke(follow_up, next_session_id)
            except subprocess.TimeoutExpired:
                return {"ok": False, "error": "经营数据已读取，但本机 Codex 汇总超过4分钟，请稍后重发。"}
            except Exception as query_answer_error:
                return {"ok": False, "error": f"经营数据已读取，但本机 Codex 汇总失败：{redact_error(str(query_answer_error))}"}

        if result.get("queries"):
            reason = "请求了已经查询过的相同数据" if repeated_query_stop else "已达到本次只读工具调用上限"
            final_instruction = (
                f"当前事件键：{event_key}。店主当前原话：{prompt}\n"
                f"外层工具循环已停止：{reason}。请只使用当前事件已经取得的事实给出最终回复，queries 必须为空。"
                "能回答多少就准确回答多少；若仍缺事实，只指出具体缺少哪一个字段和对应订单/SKU，不能泛化说数据太多或看不到表格。"
                "不要生成未经证据支持的动作。"
            )
            try:
                final_result, next_session_id = timed_invoke(final_instruction, next_session_id)
                if not final_result.get("queries"):
                    result = final_result
                else:
                    result = {
                        "reply": "我已经读取了现有经营数据，但仍缺少能唯一完成这件事的具体字段。请告诉我消息里涉及的订单号或 KW/CSX，我会继续从系统中定位。",
                        "needs_information": True,
                        "queries": [],
                        "actions": [],
                    }
            except Exception as final_error:
                return {"ok": False, "error": f"经营数据已读取，但最终整理失败：{redact_error(str(final_error))}"}

        next_turns = (
            recent_turns
            + [
                {"role": "user", "content": prompt[:1200]},
                {"role": "assistant", "content": result["reply"][:1200]},
            ]
        )[-12:]
        save_json(
            SESSION_STATE,
            {
                "session_id": "",
                "last_work_session_id_hash": hashlib.sha256(next_session_id.encode("utf-8")).hexdigest()[:16] if next_session_id else "",
                "conversation_mode": "isolated_message_session_with_explicit_recent_context",
                "recent_turns": next_turns,
                "updated_at": now_iso(),
            },
        )
        cache_response(event_key, mode, result)
        timings = {
            "lock_wait_seconds": round(lock_wait_seconds, 3),
            "model_seconds": round(model_seconds, 3),
            "query_seconds": round(query_seconds, 3),
            "query_rounds": query_round,
            "query_count": total_queries,
            "query_cache_hits": query_cache_hits,
            "runtime": "app_server" if _APP_SERVER_RUNTIME is not None else "exec_fallback",
            "total_seconds": round(time.perf_counter() - started_at, 3),
            "event_cache_hit": False,
        }
        append_performance(event_key, mode, True, timings)
        return {"ok": True, "cached": False, **result, "_timings": timings}


def _runtime_client(payload: dict[str, Any], timeout: float = 5.0) -> dict[str, Any]:
    from multiprocessing.connection import Client

    connection = Client(RUNTIME_PIPE, family="AF_PIPE", authkey=RUNTIME_AUTHKEY)
    try:
        connection.send(payload)
        if not connection.poll(timeout):
            raise TimeoutError("Codex 常驻服务响应超时")
        response = connection.recv()
        if not isinstance(response, dict):
            raise RuntimeError("Codex 常驻服务返回格式无效")
        return response
    finally:
        connection.close()


def ensure_runtime_server() -> dict[str, Any]:
    try:
        status = _runtime_client({"op": "ping"}, 1.5)
        if status.get("version") == RUNTIME_VERSION:
            return status
        try:
            _runtime_client({"op": "shutdown"}, 5)
        except Exception:
            pass
        time.sleep(0.5)
    except Exception:
        pass
    with ProcessLock(RUNTIME_START_LOCK_PATH, timeout_seconds=20):
        try:
            status = _runtime_client({"op": "ping"}, 1.5)
            if status.get("version") == RUNTIME_VERSION:
                return status
        except Exception:
            pass
        creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0) | getattr(subprocess, "DETACHED_PROCESS", 0)
        subprocess.Popen(
            [sys.executable, str(pathlib.Path(__file__).resolve()), "--serve"],
            cwd=str(PROJECT_ROOT),
            stdin=subprocess.DEVNULL,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            close_fds=True,
            creationflags=creationflags,
        )
        deadline = time.monotonic() + 30
        last_error = ""
        while time.monotonic() < deadline:
            try:
                return _runtime_client({"op": "ping"}, 2)
            except Exception as exc:
                last_error = redact_error(str(exc))
                time.sleep(0.25)
        raise RuntimeError(f"Codex 常驻服务未能启动：{last_error}")


def request_runtime(prompt: str, event_key: str, mode: str) -> dict[str, Any]:
    ensure_runtime_server()
    # A complete tool loop can legitimately take several minutes; event idempotency
    # makes a later replay safe if the local client is interrupted.
    return _runtime_client(
        {"op": "process", "prompt": prompt, "event_key": event_key, "mode": mode},
        MAX_SECONDS * (MAX_QUERY_ROUNDS + 2),
    )


def run_runtime_server() -> int:
    from multiprocessing.connection import Listener

    global _APP_SERVER_RUNTIME
    STATE_ROOT.mkdir(parents=True, exist_ok=True)
    listener = Listener(RUNTIME_PIPE, family="AF_PIPE", authkey=RUNTIME_AUTHKEY)
    _APP_SERVER_RUNTIME = CodexAppRuntime()
    save_json(
        RUNTIME_STATE_PATH,
        {"pid": os.getpid(), "status": "ready", "started_at": now_iso(), "pipe_hash": hashlib.sha256(RUNTIME_PIPE.encode()).hexdigest()[:12]},
    )
    try:
        running = True
        while running:
            connection = listener.accept()
            try:
                request = connection.recv()
                operation = str(request.get("op") or "") if isinstance(request, dict) else ""
                if operation == "ping":
                    auth_reloaded = _APP_SERVER_RUNTIME.ensure_current_auth()
                    response = {"ok": True, "runtime": "app_server", "version": RUNTIME_VERSION, "pid": os.getpid(), "query_cache_entries": len(_QUERY_CACHE), "auth_reloaded": auth_reloaded}
                elif operation == "invalidate_cache":
                    response = {"ok": True, "cleared": clear_query_cache()}
                elif operation == "shutdown":
                    response = {"ok": True, "stopping": True}
                    running = False
                elif operation == "process":
                    try:
                        response = process_message(
                            str(request.get("prompt") or ""),
                            str(request.get("event_key") or ""),
                            str(request.get("mode") or "normal"),
                        )
                    except Exception as exc:
                        response = {"ok": False, "error": f"本机 Codex 桥接失败：{redact_error(str(exc))}"}
                else:
                    response = {"ok": False, "error": "未知的 Codex 常驻服务操作"}
                connection.send(response)
                save_json(RUNTIME_STATE_PATH, {"pid": os.getpid(), "status": "ready", "updated_at": now_iso()})
            except (EOFError, OSError):
                pass
            finally:
                connection.close()
    finally:
        listener.close()
        if _APP_SERVER_RUNTIME is not None:
            _APP_SERVER_RUNTIME.close()
        save_json(RUNTIME_STATE_PATH, {"pid": os.getpid(), "status": "stopped", "updated_at": now_iso()})
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="飞书私聊到本机 Codex 的安全桥接器")
    parser.add_argument("--check", action="store_true")
    parser.add_argument("--serve", action="store_true")
    parser.add_argument("--ensure-runtime", action="store_true")
    parser.add_argument("--invalidate-cache", action="store_true")
    parser.add_argument("--prompt-base64")
    parser.add_argument("--event-key")
    parser.add_argument("--mode", choices=["normal", "preview", "execute"], default="normal")
    return parser


def main() -> int:
    args = build_parser().parse_args()
    if args.serve:
        return run_runtime_server()
    if args.ensure_runtime:
        try:
            return output(ensure_runtime_server())
        except Exception as exc:
            return output({"ok": False, "error": redact_error(str(exc))})
    if args.check:
        status = check_runtime()
        try:
            status["persistent_runtime"] = _runtime_client({"op": "ping"}, 1.5)
        except Exception:
            status["persistent_runtime"] = {"ok": False, "status": "stopped"}
        return output(status)
    if args.invalidate_cache:
        try:
            ensure_runtime_server()
            return output(_runtime_client({"op": "invalidate_cache"}, 5))
        except Exception as exc:
            return output({"ok": False, "error": redact_error(str(exc))})
    if not args.prompt_base64 or not args.event_key:
        return output({"ok": False, "error": "缺少消息内容或事件标识"})
    try:
        prompt = decode_prompt(args.prompt_base64)
    except ValueError as exc:
        return output({"ok": False, "error": str(exc)})
    if not prompt:
        return output({"ok": False, "error": "消息内容为空"})
    try:
        try:
            return output(request_runtime(prompt, args.event_key, args.mode))
        except Exception:
            # The established one-shot path remains available if the local daemon
            # cannot start, so responsiveness work never disables Feishu chat.
            return output(process_message(prompt, args.event_key, args.mode))
    except TimeoutError as exc:
        return output({"ok": False, "error": str(exc)})
    except Exception as exc:
        return output({"ok": False, "error": f"本机 Codex 桥接失败：{redact_error(str(exc))}"})


if __name__ == "__main__":
    raise SystemExit(main())
