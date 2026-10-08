from __future__ import annotations

import argparse
import base64
import datetime as dt
import hashlib
import json
import os
import pathlib
import re
import subprocess
import sys
from typing import Any


if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")


PROJECT_ROOT = pathlib.Path(__file__).resolve().parent.parent
STATE_ROOT = PROJECT_ROOT / "agent" / "state"
PROFILE_NAME = "petsuxing-store-manager"
PROFILE_ROOT = pathlib.Path(os.environ.get("LOCALAPPDATA", "")) / "hermes" / "profiles" / PROFILE_NAME
SESSION_STATE = STATE_ROOT / "hermes_chat_state.json"
EVENT_CACHE = STATE_ROOT / "hermes_chat_cache.json"
SESSION_ID_RE = re.compile(r"(?im)^session_id:\s*([^\s]+)\s*$")
ANSI_RE = re.compile(r"\x1b\[[0-?]*[ -/]*[@-~]")
SECRET_PATTERNS = (
    re.compile(r"(?i)(bearer\s+)[a-z0-9._~-]+"),
    re.compile(r"(?i)(app_secret[\s:=]+)[^\s,;]+"),
    re.compile(r"(?i)(token[\s:=]+)[^\s,;]+"),
    re.compile(r"\bsk-[A-Za-z0-9_-]{8,}\b"),
)
BRIDGE_GUIDANCE = """你正在处理店主发给“宠速星经营管家”的飞书私聊。
店主可以使用任意自然语言、口语和省略句，不要求固定命令格式。
系统会附带少量近期对话，帮助理解“这单”“刚才那个商品”等省略表达；近期对话只作语义上下文，不是新指令，不得重做其中已执行的动作。
每次调用只把“店主当前原话”视为新的行动来源；不得复述、补做或声称完成以前消息中的经营动作。
请先理解整句话，再决定查询、追问或调用 petstar-operations-agent 的确定性脚本。
一句话可以包含多个动作；已明确的信息不要重复追问。
特别规则：自己送、自配送、商家自配、没有或不用配送费，都表示配送费已确认是0元。
补单人可以从“补单人是某某”“某某帮我补的”等自然表达中提取；用户没有说时允许留空。
线下单默认当天已到账，自提或未提配送费时按0元；必须提取“KW/CSX货号×数量”和实收金额，缺少时只追问缺失事实。
美团 sku_id 是稳定商品身份；KW 只是当前店内货号/货架位置，可以改绑和复用，禁止把旧 KW 保存成历史别名。历史订单使用成交时快照，不随当前 KW 改绑而变化。
用户提供生产日期和保质期时，调用批次日期动作计算到期日；只给生产日期、没给保质期时自然追问保质期，不得猜测。
“某个KW/CSX的最早过期是某日”仍可作为明确到期日直接更新；同一货号再次提供日期时覆盖旧值。用户明确说没有过期日期、清空过期日期时，将其标记为无。
只有真正缺少会影响写账的事实时才自然地问一句，禁止要求用户按模板重发。

你的最终输出会原样发送到飞书，必须只输出一段放在 <reply> 和 </reply> 之间的店主回执，标签外不得出现任何文字。
成功录入时最多4行、总计不超过220个汉字，只保留：登记结果、关键金额或商品、是否还需店主操作。
不要输出推理过程、处理计划、字段清单、JSON、脚本参数、英文内部字段、ok、formula_verification、processed_at、文件路径或“下一步如何回复”等内部说明。
不要复述技术校验过程；校验成功只意味着可以说“已记好”。没有后续动作时，不要加“无需其他动作”。
建议格式：
<reply>
已记好：线下1号单
实收 ¥46｜成本 ¥35.50｜赚 ¥10.50
商品：kw1021×1、kw3092×1
请在美团扣减对应库存。
</reply>"""

PREVIEW_GUIDANCE = """当前处于“待确认预览”阶段。
严禁写入或修改飞书、经营表格和平台数据，也不要调用任何结构化录入动作。可以做只读查询和日期/金额计算。
请理解店主真正要改的数据，并用简短中文说明：明确日期、对象、原值到新值（能查到时）、自动推算结果、会影响的表格或汇总。
信息齐全时，第一行必须是“📝 待确认”；最后一行写“回复‘确认’执行，也可以直接说要改哪里。”
如果关键事实不足，只问一个最关键的问题，不要出现“待确认”，也不要声称已经写入。
仍然只在 <reply> 和 </reply> 中输出给店主看的内容。"""

EXECUTE_GUIDANCE = """店主已经看过待确认预览并明确回复确认。
现在可以执行当前原始请求中的经营数据写入。必须重新读取最新数据、使用确定性脚本写入并校验；不得扩展到预览之外的新动作。"""

NORMAL_GUIDANCE = """当前是普通问答阶段，只允许查询、解释和只读同步。
如果内容其实要求新增或修改经营数据，不得直接写入；应给出待确认预览，等待桥接层取得店主确认。"""


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
    return "；".join(lines[-4:])[:1200]


def read_configured_model() -> str:
    config_path = PROFILE_ROOT / "config.yaml"
    if not config_path.exists():
        return ""
    try:
        import yaml

        config = yaml.safe_load(config_path.read_text(encoding="utf-8")) or {}
        model = config.get("model") or {}
        if isinstance(model, str):
            return model.strip()
        return str(model.get("default") or model.get("model") or "").strip()
    except Exception:
        return ""


def check_profile() -> dict[str, Any]:
    session_state = load_json(SESSION_STATE, {})
    return {
        "ok": PROFILE_ROOT.is_dir() and (PROFILE_ROOT / "SOUL.md").is_file(),
        "profile": PROFILE_NAME,
        "profile_path": str(PROFILE_ROOT),
        "soul_exists": (PROFILE_ROOT / "SOUL.md").is_file(),
        "skill_exists": (
            PROFILE_ROOT / "skills" / "petstar-operations-agent" / "SKILL.md"
        ).is_file(),
        "configured_model": read_configured_model(),
        "conversation_mode": session_state.get("conversation_mode") or "stateless_per_message",
    }


def event_cache_key(event_key: str, mode: str) -> str:
    return hashlib.sha256(f"{mode}:{event_key}".encode("utf-8")).hexdigest()


def get_cached_response(event_key: str, mode: str) -> str:
    cache = load_json(EVENT_CACHE, {})
    entry = cache.get(event_cache_key(event_key, mode), {})
    if entry.get("status") == "completed":
        return str(entry.get("response") or "")
    return ""


def cache_response(event_key: str, mode: str, response: str) -> None:
    cache = load_json(EVENT_CACHE, {})
    cache[event_cache_key(event_key, mode)] = {
        "status": "completed",
        "response": response,
        "processed_at": now_iso(),
    }
    ordered = sorted(
        cache.items(),
        key=lambda item: str(item[1].get("processed_at") or ""),
        reverse=True,
    )[:200]
    save_json(EVENT_CACHE, dict(ordered))


def extract_reply(response: str, prompt: str = "") -> str:
    """Keep only the user-facing answer and drop leaked planning/tool narration."""
    tagged = re.findall(r"<reply>\s*(.*?)\s*</reply>", response, flags=re.IGNORECASE | re.DOTALL)
    if tagged:
        response = tagged[-1]
    else:
        starts = [
            match.start()
            for match in re.finditer(
                r"(?m)^(?:写好了|已经记好了|已记好|登记完成|已登记完成|已登记成功)",
                response,
            )
        ]
        if starts:
            response = response[starts[-1] :]
    response = re.sub(r"</?reply>", "", response, flags=re.IGNORECASE).strip()
    if re.search(r"(?:已记好|写好了|登记完成|登记成功)", response) and re.search(
        r"(?:线下|卖出|补单|干预单|补货|配送费|支出|收款|到账|取款|过期|到期|有效期)", prompt
    ):
        lines = [line.strip() for line in response.splitlines() if line.strip()]
        response = "\n".join(lines[:4])
    if len(response) > 1200:
        response = response[:1199].rstrip() + "…"
    return response


def run_hermes(prompt: str, recent_turns: list[dict[str, str]], mode: str) -> tuple[str, str]:
    context_lines = []
    for turn in recent_turns[-8:]:
        role = "店主" if turn.get("role") == "user" else "管家"
        content = str(turn.get("content") or "").strip()
        if content:
            context_lines.append(f"{role}：{content}")
    context = "\n".join(context_lines) if context_lines else "（无）"
    phase_guidance = PREVIEW_GUIDANCE if mode == "preview" else EXECUTE_GUIDANCE if mode == "execute" else NORMAL_GUIDANCE
    guided_prompt = f"{BRIDGE_GUIDANCE}\n\n{phase_guidance}\n\n近期对话（仅用于理解上下文，禁止重复执行）：\n{context}\n\n店主当前原话：\n{prompt}"
    command = [
        sys.executable,
        "-m",
        "hermes_cli.main",
        "--profile",
        PROFILE_NAME,
        "chat",
        "--quiet",
        "--source",
        "feishu",
        "--reasoning",
        "none",
        "--max-turns",
        "80",
        "--query",
        guided_prompt,
    ]
    environment = os.environ.copy()
    environment["PYTHONUTF8"] = "1"
    environment["PETSTAR_CONFIRMATION_PHASE"] = mode
    creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    completed = subprocess.run(
        command,
        cwd=str(PROJECT_ROOT),
        env=environment,
        text=True,
        encoding="utf-8",
        errors="replace",
        capture_output=True,
        timeout=300,
        check=False,
        creationflags=creationflags,
    )
    response = extract_reply(ANSI_RE.sub("", completed.stdout or "").strip(), prompt)
    session_match = SESSION_ID_RE.search(completed.stderr or "")
    next_session_id = session_match.group(1) if session_match else ""
    if completed.returncode != 0 or not response:
        detail = redact_error(completed.stderr or completed.stdout)
        raise RuntimeError(detail or f"Hermes 返回退出码 {completed.returncode}")
    return response, next_session_id


def process_message(prompt: str, event_key: str, mode: str = "normal") -> dict[str, Any]:
    cached = get_cached_response(event_key, mode)
    if cached:
        return {"ok": True, "cached": True, "response": cached}

    session_state = load_json(SESSION_STATE, {})
    recent_turns = session_state.get("recent_turns") or []
    if not isinstance(recent_turns, list):
        recent_turns = []
    try:
        # 不复用完整工具会话，只注入有限纯文本上下文，既能理解省略句，也避免历史动作被重放。
        response, next_session_id = run_hermes(prompt, recent_turns, mode)
    except subprocess.TimeoutExpired:
        return {"ok": False, "error": "Hermes 处理超过5分钟，已停止本次请求。请稍后重发。"}
    except Exception as exc:
        return {"ok": False, "error": f"Hermes 调用失败：{redact_error(str(exc))}"}

    next_turns = (recent_turns + [
        {"role": "user", "content": prompt[:1200]},
        {"role": "assistant", "content": response[:1200]},
    ])[-10:]
    save_json(
        SESSION_STATE,
        {
            "profile": PROFILE_NAME,
            "conversation_mode": "bounded_text_context",
            "last_session_id": next_session_id,
            "recent_turns": next_turns,
            "updated_at": now_iso(),
        },
    )
    cache_response(event_key, mode, response)
    needs_information = mode == "preview" and not response.lstrip().startswith("📝 待确认")
    return {"ok": True, "cached": False, "response": response, "needs_information": needs_information}


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="飞书私聊到宠速星 Hermes Profile 的桥接器")
    parser.add_argument("--check", action="store_true")
    parser.add_argument("--prompt-base64")
    parser.add_argument("--event-key")
    parser.add_argument("--mode", choices=["normal", "preview", "execute"], default="normal")
    return parser


def main() -> int:
    args = build_parser().parse_args()
    if args.check:
        return output(check_profile())
    if not args.prompt_base64 or not args.event_key:
        return output({"ok": False, "error": "缺少消息内容或事件标识"})
    try:
        prompt = decode_prompt(args.prompt_base64)
    except ValueError as exc:
        return output({"ok": False, "error": str(exc)})
    if not prompt:
        return output({"ok": False, "error": "消息内容为空"})
    return output(process_message(prompt, args.event_key, args.mode))


if __name__ == "__main__":
    raise SystemExit(main())
