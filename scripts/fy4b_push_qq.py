# -*- coding: utf-8 -*-
"""
FY4B QQ 直接推送脚本（QClaw 备用方案）

功能：
  1. 检测是否由 QClaw 调用 —— 通过 OPENCLAW_SESSION_ID 环境变量判断
  2. 若由 QClaw 调用（变量存在），说明 QClaw 推送逻辑会接管，直接退出
  3. 若非 QClaw 调用，则从 last_output.txt 读取 fy4b_download.py 的输出，
     解析 Markdown 报告，并通过 QQ 开放平台 API 直接推送到 QQ

依赖：requests 库（Windows 系统自带或 pip install requests）
"""

import json, pathlib, sys, re, os, subprocess

# ============ 配置读取 ============
_SCRIPT_DIR = pathlib.Path(__file__).parent
_config_path = _SCRIPT_DIR / "skill_config.json"
_config = json.loads(_config_path.read_text(encoding="utf-8"))

# ============ 步骤 1：检测是否由 QClaw 调用 ============
def is_run_by_qclaw():
    """OPENCLAW_SESSION_ID 存在 = 由 QClaw cron agentTurn 调用"""
    return bool(os.environ.get("OPENCLAW_SESSION_ID"))

# ============ 步骤 2：从 last_output.txt 解析报告 ============
def parse_report_from_output(raw: str):
    """从 fy4b_download.py 的原始输出中提取 Markdown 报告"""
    # 匹配 📡 **FY4B 云图下载报告** 开始到文件末尾（或下一个 [DONE] 标记前）
    match = re.search(
        r"(📡\s+\*\*FY4B[^\n]*\*\*[\s\S]*)",
        raw
    )
    if match:
        return match.group(1).strip()
    # 回退：取末尾 1000 字
    return raw[-1000:].strip()

# ============ 步骤 3：通过 QQ 开放平台 API 直接推送 ============
def get_qqbot_access_token():
    """读取 QQ 开放平台 access_token"""
    token_cache = _SCRIPT_DIR / "qqbot_token.txt"
    if not token_cache.exists():
        return None
    return token_cache.read_text(encoding="utf-8").strip()

def _do_push(message: str, to_user: str, token: str) -> bool:
    """
    通过 QQ 机器人 API 直接推送私聊消息（C2C，主动消息）

    官方文档: https://api.bot.qq.com/v2/users/{user_openid}/messages
    鉴权:     Authorization: QQBot <token>
    Body:     { "msg_type": 0, "content": "<text>" }
    """
    import urllib.request, json as _json

    url = f"https://api.bot.qq.com/v2/users/{to_user}/messages"
    payload = _json.dumps({
        "msg_type": 0,
        "content": message
    }, ensure_ascii=False).encode("utf-8")

    req = urllib.request.Request(
        url,
        data=payload,
        headers={
            "Authorization": f"QQBot {token}",
            "Content-Type": "application/json; charset=utf-8",
        },
        method="POST"
    )

    with urllib.request.urlopen(req, timeout=15) as resp:
        result = _json.loads(resp.read().decode("utf-8"))
        # 成功时响应包含 id 字段
        if result.get("id"):
            return True
        print(f"[QQPush] API 返回错误: {result}", file=sys.stderr)
        return False

def push_to_qq(message: str, to_user: str, app_id: str, token: str) -> bool:
    try:
        return _do_push(message, to_user, token)
    except Exception as e:
        print(f"[QQPush] urllib 推送失败: {e}", file=sys.stderr)
        return False

def push_to_qq_via_curl(message: str, to_user: str, app_id: str, token: str) -> bool:
    """PowerShell Invoke-WebRequest 备选推送（urllib 不可用时）"""
    import json as _json, subprocess as _sub, base64 as _b64
    url = f"https://api.bot.qq.com/v2/users/{to_user}/messages"
    payload_bytes = _json.dumps({
        "msg_type": 0,
        "content": message
    }, ensure_ascii=False).encode("utf-8")
    b64 = _b64.b64encode(payload_bytes).decode("ascii")
    cmd = [
        "powershell", "-Command",
        f"$body = [System.Convert]::FromBase64String('{b64}'); "
        f"$resp = Invoke-WebRequest -Uri '{url}' -Method POST "
        f"-Headers @{{'Authorization'='QQBot {token}'; 'Content-Type'='application/json; charset=utf-8'}} "
        f"-Body $body -TimeoutSec 15; $resp.Content"
    ]
    try:
        result = _sub.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")
        data = _json.loads(result.stdout.strip())
        if data.get("id"):
            return True
        print(f"[QQPush] curl 方式 API 返回错误: {data}", file=sys.stderr)
        return False
    except Exception as e:
        print(f"[QQPush] curl 推送失败: {e}", file=sys.stderr)
        return False

def main():
    # ── 步骤 1：检测 QClaw ──
    if is_run_by_qclaw():
        print("[QQPush] QClaw 环境检测到，跳过直接推送（由 QClaw message 工具接管）")
        sys.exit(0)

    print("[QQPush] 非 QClaw 环境，准备直接推送 QQ")

    # ── 步骤 2：读取 fy4b_download.py 输出 ──
    output_file = _SCRIPT_DIR / "last_output.txt"
    if not output_file.exists():
        print("[QQPush] last_output.txt 不存在，无法推送", file=sys.stderr)
        sys.exit(1)

    raw_output = output_file.read_text(encoding="utf-8", errors="replace")
    report = parse_report_from_output(raw_output)
    if not report or "⚠️" in report and "FY4B 云图下载报告" not in report:
        print("[QQPush] 未找到有效报告，跳过推送", file=sys.stderr)
        print(f"[QQPush] 原始内容末尾: {raw_output[-300:]}", file=sys.stderr)
        sys.exit(0)

    print(f"[QQPush] 报告内容：\n{report[:200]}...")

    # ── 步骤 3：推送 QQ ──
    # 从配置文件读取 QQ 推送目标
    qq_target = None
    for t in _config.get("push_targets", []):
        if t.get("channel") == "qqbot":
            qq_target = t
            break

    if not qq_target:
        print("[QQPush] 配置文件中未找到 qqbot 推送目标，跳过", file=sys.stderr)
        sys.exit(0)

    to_user = qq_target["to"].replace("qqbot:c2c:", "")
    app_id = qq_target.get("accountId", "")
    token = get_qqbot_access_token()

    if not token:
        print("[QQPush] 未找到 QQ access_token（qqbot_token.txt），跳过推送", file=sys.stderr)
        print("[QQPush] 提示：请确保 QQ 开放平台机器人已授权，token 存放于 scripts/qqbot_token.txt", file=sys.stderr)
        sys.exit(0)

    # 优先用 urllib，失败则回退 curl
    success = push_to_qq(message=report, to_user=to_user, app_id=app_id, token=token)
    if not success:
        # 回退 curl
        print("[QQPush] urllib 方式失败，尝试 PowerShell Invoke-WebRequest")
        success = push_to_qq_via_curl(message=report, to_user=to_user, app_id=app_id, token=token)

    if success:
        print("[QQPush] ✅ QQ 推送成功")
    else:
        print("[QQPush] ❌ QQ 推送失败，请检查 token 和 openid", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    main()
