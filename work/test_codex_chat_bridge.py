from __future__ import annotations

import importlib.util
import pathlib
import tempfile


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

runtime = bridge.check_runtime()
assert runtime["query_tool_exists"], "经营数据只读查询工具未就绪"

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
