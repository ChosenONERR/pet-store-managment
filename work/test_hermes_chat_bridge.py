from __future__ import annotations

import importlib.util
import pathlib
import tempfile


PROJECT_ROOT = pathlib.Path(__file__).resolve().parent.parent
MODULE_PATH = PROJECT_ROOT / "agent" / "hermes_chat_bridge.py"
spec = importlib.util.spec_from_file_location("petstar_hermes_chat_bridge", MODULE_PATH)
assert spec and spec.loader
bridge = importlib.util.module_from_spec(spec)
spec.loader.exec_module(bridge)


with tempfile.TemporaryDirectory() as temp_dir:
    temp_root = pathlib.Path(temp_dir)
    bridge.SESSION_STATE = temp_root / "session.json"
    bridge.EVENT_CACHE = temp_root / "cache.json"
    calls: list[tuple[str, str]] = []

    def fake_run_hermes(
        prompt: str, recent_turns: list[dict[str, str]], mode: str
    ) -> tuple[str, str]:
        calls.append((prompt, mode))
        response = f"📝 待确认：{prompt}" if mode == "preview" else f"收到：{prompt}"
        return response, f"session-{len(calls)}"

    bridge.run_hermes = fake_run_hermes
    first = bridge.process_message("上一条经营消息", "event-1")
    second = bridge.process_message("我将 kw1003 修改成 kw1106", "event-2")
    preview = bridge.process_message("昨天配送费改成8元", "event-3", "preview")
    duplicate = bridge.process_message("不应重新执行", "event-2")
    state = bridge.load_json(bridge.SESSION_STATE, {})

    assert first["ok"] and second["ok"] and preview["ok"]
    assert len(calls) == 3, calls
    assert calls[1] == ("我将 kw1003 修改成 kw1106", "normal")
    assert calls[2] == ("昨天配送费改成8元", "preview")
    assert preview["needs_information"] is False
    assert duplicate["cached"] is True
    assert state["conversation_mode"] == "bounded_text_context"

print("Hermes message isolation OK")
