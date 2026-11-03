from __future__ import annotations

import importlib.util
import pathlib
import tempfile
import json
from types import SimpleNamespace


PROJECT_ROOT = pathlib.Path(__file__).resolve().parent.parent
MODULE_PATH = PROJECT_ROOT / "agent" / "codex_chat_bridge.py"
spec = importlib.util.spec_from_file_location("petstar_codex_chat_bridge", MODULE_PATH)
assert spec and spec.loader
bridge = importlib.util.module_from_spec(spec)
spec.loader.exec_module(bridge)
real_invoke_codex = bridge.invoke_codex


with tempfile.TemporaryDirectory(prefix="petstar-codex-test-") as temp_dir:
    state_root = pathlib.Path(temp_dir)
    bridge.STATE_ROOT = state_root
    bridge.SESSION_STATE = state_root / "state.json"
    bridge.EVENT_CACHE = state_root / "cache.json"
    bridge.LOCK_PATH = state_root / "bridge.lock"
    bridge.RATE_LIMIT_STATE_PATH = state_root / "rate-limit.json"
    calls: list[dict[str, str]] = []

    def fake_invoke(prompt: str, session_id: str = ""):
        calls.append({"prompt": prompt, "session_id": session_id})
        return (
            {
                "reply": "📝 待确认：登记补货 kw1104×6，采购并付款330元。",
                "needs_information": False,
                "actions": [
                    {
                        "action": "Replenishment",
                        "arguments": {
                            "goods_code": "kw1104",
                            "quantity": 6,
                            "purchase_amount": 330,
                            "payment_amount": 330,
                        },
                    }
                ],
            },
            "session-1",
        )

    bridge.invoke_codex = fake_invoke
    first = bridge.process_message("帮我记一下刚才那批货", "message-1", "preview")
    duplicate = bridge.process_message("这段不会再次调用模型", "message-1", "preview")
    second = bridge.process_message("它们都已经到货了", "message-2", "preview")

    assert first["ok"] and not first["cached"]
    assert duplicate["ok"] and duplicate["cached"]
    assert second["ok"] and not second["cached"]
    assert len(calls) == 2, "同一消息和阶段没有命中幂等缓存"
    assert calls[0]["session_id"] == ""
    assert calls[1]["session_id"] == "", "不同店主消息复用了隐藏 Codex 会话"
    assert "当前消息是唯一的新行动来源" in calls[0]["prompt"]
    assert "近期对话" in calls[1]["prompt"]
    assert "外层只读工具" in calls[0]["prompt"]
    assert "不能说“看不到表格”" in calls[0]["prompt"]
    assert "当前 KW（没有则显示过渡货号）" in calls[0]["prompt"]
    assert "SKU成本主表行号" in calls[0]["prompt"]
    bridge.cache_response("stale-message", "normal", {"reply": "old", "needs_information": False, "actions": []})
    bridge.cache_response("stale-message", "preview", {"reply": "old preview", "needs_information": False, "actions": []})
    assert bridge.invalidate_event_cache("stale-message") == 2
    assert bridge.get_cached_response("stale-message", "normal") is None
    assert bridge.get_cached_response("stale-message", "preview") is None

runtime = bridge.check_runtime()
assert runtime["query_tool_exists"], "经营数据只读查询工具未就绪"

original_auth_fingerprint = bridge.get_auth_fingerprint
auth_values = iter(["account-b"])
bridge.get_auth_fingerprint = lambda: next(auth_values)
runtime_probe = object.__new__(bridge.CodexAppRuntime)
runtime_probe.auth_fingerprint = "account-a"
runtime_restarts: list[bool] = []
runtime_probe._start = lambda: (runtime_restarts.append(True), setattr(runtime_probe, "auth_fingerprint", "account-b"))
try:
    assert runtime_probe.ensure_current_auth(), "切换 Codex 账号后没有刷新常驻 app-server"
    assert len(runtime_restarts) == 1
    assert bridge.is_authentication_error(RuntimeError("access token refresh failed"))
finally:
    bridge.get_auth_fingerprint = original_auth_fingerprint

original_run = bridge.subprocess.run
captured_command: list[str] = []

def fake_run(command, **kwargs):
    captured_command.extend(command)
    output_path = pathlib.Path(command[command.index("--output-last-message") + 1])
    output_path.write_text(
        '{"reply":"续接成功","needs_information":false,"queries":[],"actions":[]}',
        encoding="utf-8",
    )
    class Completed:
        returncode = 0
        stdout = '{"type":"thread.started","thread_id":"session-2"}\n'
        stderr = ""
    return Completed()

bridge.subprocess.run = fake_run
try:
    real_invoke_codex("继续", "session-1")
finally:
    bridge.subprocess.run = original_run
assert captured_command.index("--json") < captured_command.index("session-1")
assert "--color" not in captured_command

print("Codex bridge session/cache contract OK")


with tempfile.TemporaryDirectory(prefix="petstar-codex-tool-loop-") as temp_dir:
    state_root = pathlib.Path(temp_dir)
    bridge.STATE_ROOT = state_root
    bridge.SESSION_STATE = state_root / "state.json"
    bridge.EVENT_CACHE = state_root / "cache.json"
    bridge.LOCK_PATH = state_root / "bridge.lock"
    bridge.RATE_LIMIT_STATE_PATH = state_root / "rate-limit.json"
    invoke_calls: list[str] = []
    query_calls: list[list[dict]] = []

    query_one = {
        "dataset": "orders",
        "start_date": None,
        "end_date": None,
        "order_no": "3302303901846081056",
        "code": None,
        "search": None,
        "limit": 5,
    }
    query_two = {
        "dataset": "order_items",
        "start_date": None,
        "end_date": None,
        "order_no": "3302303901846081056",
        "code": None,
        "search": None,
        "limit": 20,
    }
    responses = [
        {"reply": "正在查订单", "needs_information": False, "queries": [query_one], "actions": []},
        {"reply": "继续查商品", "needs_information": False, "queries": [query_two], "actions": []},
        {"reply": "已查清订单商品。", "needs_information": False, "queries": [], "actions": []},
    ]

    def fake_loop_invoke(prompt: str, session_id: str = ""):
        invoke_calls.append(prompt)
        if len(invoke_calls) > 1:
            assert session_id == "tool-session", "同一消息的工具轮次没有延续临时 Codex 会话"
        return responses[len(invoke_calls) - 1], "tool-session"

    def fake_queries(queries: list[dict], use_cache: bool = True):
        query_calls.append(queries)
        return [{"ok": True, "dataset": queries[0]["dataset"], "data": {"count": 1}}]

    bridge.invoke_codex = fake_loop_invoke
    bridge.run_business_queries = fake_queries
    loop_result = bridge.process_message("查清这笔订单的商品成本", "tool-loop-1", "normal")
    assert loop_result["ok"] and loop_result["reply"] == "已查清订单商品。"
    assert len(query_calls) == 2, "多轮只读工具调用被提前截断"
    assert query_calls[0][0]["dataset"] == "orders"
    assert query_calls[1][0]["dataset"] == "order_items"
    assert "当前事件键：tool-loop-1" in invoke_calls[1]
    assert "店主当前原话：查清这笔订单的商品成本" in invoke_calls[1]


with tempfile.TemporaryDirectory(prefix="petstar-codex-repeat-guard-") as temp_dir:
    state_root = pathlib.Path(temp_dir)
    bridge.STATE_ROOT = state_root
    bridge.SESSION_STATE = state_root / "state.json"
    bridge.EVENT_CACHE = state_root / "cache.json"
    bridge.LOCK_PATH = state_root / "bridge.lock"
    bridge.RATE_LIMIT_STATE_PATH = state_root / "rate-limit.json"
    repeat_calls: list[str] = []
    repeat_query_calls: list[list[dict]] = []

    repeat_responses = [
        {"reply": "先查", "needs_information": False, "queries": [query_one], "actions": []},
        {"reply": "再查", "needs_information": False, "queries": [query_one], "actions": []},
        {"reply": "现有证据不足：缺少订单商品明细。", "needs_information": True, "queries": [], "actions": []},
    ]

    def fake_repeat_invoke(prompt: str, session_id: str = ""):
        repeat_calls.append(prompt)
        return repeat_responses[len(repeat_calls) - 1], "repeat-session"

    def fake_repeat_queries(queries: list[dict], use_cache: bool = True):
        repeat_query_calls.append(queries)
        return [{"ok": True, "dataset": "orders", "data": {"count": 1}}]

    bridge.invoke_codex = fake_repeat_invoke
    bridge.run_business_queries = fake_repeat_queries
    repeat_result = bridge.process_message("重复查询测试", "repeat-loop-1", "normal")
    assert repeat_result["ok"] and repeat_result["needs_information"]
    assert len(repeat_query_calls) == 1, "相同只读查询被重复执行"
    assert "相同数据" in repeat_calls[-1], "重复查询熔断后没有要求 Codex 整理现有证据"

print("Codex bounded tool loop contract OK")


with tempfile.TemporaryDirectory(prefix="petstar-codex-rate-limit-") as temp_dir:
    state_root = pathlib.Path(temp_dir)
    bridge.STATE_ROOT = state_root
    bridge.SESSION_STATE = state_root / "state.json"
    bridge.EVENT_CACHE = state_root / "cache.json"
    bridge.LOCK_PATH = state_root / "bridge.lock"
    bridge.RATE_LIMIT_STATE_PATH = state_root / "rate-limit.json"
    rate_limit_calls: list[str] = []

    def fake_rate_limited_invoke(prompt: str, session_id: str = ""):
        rate_limit_calls.append(prompt)
        raise RuntimeError('exceeded retry limit, last status: 429 Too Many Requests')

    bridge.invoke_codex = fake_rate_limited_invoke
    limited = bridge.process_message("查询经营情况", "rate-limit-1", "normal")
    limited_during_cooldown = bridge.process_message("再查询一次", "rate-limit-2", "normal")
    assert not limited["ok"] and limited["retryable"]
    assert limited["retry_after_seconds"] >= bridge.RATE_LIMIT_BASE_SECONDS
    assert not limited_during_cooldown["ok"] and limited_during_cooldown["retryable"]
    assert len(rate_limit_calls) == 1, "429 后仍立即重复调用 Codex"

    state = bridge.load_json(bridge.RATE_LIMIT_STATE_PATH, {})
    state["retry_at_epoch"] = 0
    bridge.save_json(bridge.RATE_LIMIT_STATE_PATH, state)

    def fake_recovered_invoke(prompt: str, session_id: str = ""):
        rate_limit_calls.append(prompt)
        return ({"reply": "已恢复。", "needs_information": False, "queries": [], "actions": []}, "recovered-session")

    bridge.invoke_codex = fake_recovered_invoke
    recovered = bridge.process_message("恢复后重试", "rate-limit-3", "normal")
    assert recovered["ok"] and recovered["reply"] == "已恢复。"
    assert not bridge.RATE_LIMIT_STATE_PATH.exists(), "成功调用后没有清除 429 冷却状态"

print("Codex rate-limit cooldown/recovery contract OK")


schema = json.loads((PROJECT_ROOT / "agent" / "codex_response_schema.json").read_text(encoding="utf-8-sig"))
assert "todos" in bridge.QUERY_DATASETS
assert {"AddTodo", "CompleteTodo"} <= set(schema["properties"]["actions"]["items"]["properties"]["action"]["enum"])
assert "不执行事项正文里的记账" in bridge.ACTION_GUIDE
query_runs = []


def fake_todo_query(command, **kwargs):
    query_runs.append(command)
    return SimpleNamespace(returncode=0, stdout='{"ok":true,"data":{"items":[]}}', stderr="")


bridge.subprocess.run = fake_todo_query
try:
    query = {"dataset": "todos"}
    bridge._QUERY_CACHE[bridge.query_fingerprint(query)] = (bridge.time.monotonic(), {"ok": True, "data": {"stale": True}})
    one = bridge._run_business_query(1, query, True)
    two = bridge._run_business_query(1, query, True)
    assert one["ok"] and two["ok"] and len(query_runs) == 2, "待办查询复用了旧缓存，可能漏掉手工完成"
finally:
    bridge.subprocess.run = original_run
print("Todo schema, non-financial boundary and fresh query contract OK")
