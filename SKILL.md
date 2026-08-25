---
name: FY4B-auto-download
description: FY4B 卫星云图自动下载技能。根据定时触发，自动检查待导入批次、通过 Python 脚本下载 FY4B 云图、更新 Excel 过期时间并生成新的下载链接列表。触发词包括：FY4B云图、定时下载卫星云图、FY4B自动下载、卫星云图监控。当用户提到 FY4B 下载任务、定时任务配置、云图自动导入时使用此技能。
---

# FY4B 云图自动下载

## 配置读取

所有关键配置从 `scripts/skill_config.json` 读取（相对于 skill 目录）。**不要硬编码任何路径或账号信息**，一律从配置文件读取。

> ⚠️ 配置文件路径：`{skill_dir}/scripts/skill_config.json`

读取配置示例：
```python
import json, pathlib
_config_path = pathlib.Path(__file__).parent / "skill_config.json"
_config = json.loads(_config_path.read_text(encoding="utf-8"))
```

### 配置字段说明

| 字段 | 说明 |
|------|------|
| `paths.staging_dir` | 临时下载目录（STAGING），下载完成后移动至目标目录 |
| `paths.target_dir` | 最终图片保存目录 |
| `paths.urls_txt_file` | 截止T.txt 文件所在目录 |
| `paths.Excel_file` | Excel 文件路径（含链接过期时间） |
| `paths.url_template_file` | URL 模板文件（相对路径，`templates/url_template.md`） |
| `paths.idm_program` | IDM 程序路径（IDM 导入方案备选） |
| `paths.downloader_script` | 核心下载脚本文件名（`fy4b_direct_download.py`） |
| `push_targets` | 推送目标列表，每项含 `channel`、`to`、`accountId` |
| `push_message_template` | 推送消息 Markdown 模板 |
| `cron_schedule` | cron 表达式，固定 `30 9,15,21 * * *` |

## 工作流程

### 定时触发（agentTurn）

由 OpenClaw cron 定时触发，payload 指示 agent 执行以下步骤：

**步骤1：执行 FY4B 云图下载脚本获取输出。**

运行以下命令并捕获完整输出：
```
python {skills_dir}\FY4B-auto-download\scripts\fy4b_download.py
```

**步骤2：从输出中提取 Markdown 报告。**

- 如果输出包含「📡 **FY4B 云图下载报告**」则为成功，直接使用该 Markdown 内容。
- 如果输出包含「⚠️ 解析失败」则为失败，使用错误原文作为消息。

**步骤3：用 message 工具推送到以下目标（依次调用，channel/to/accountId 必须严格匹配）：**
- 目标1: `channel="qqbot", to="qqbot:c2c:<your_qq_openid>", accountId="<qqbot_AppID>"`
- 目标2: `channel="openclaw-weixin", to="<chat_id>@im.wechat", accountId="<openclaw_weixin_accountId>"`

> 💡 首次设置定时任务时：读取配置文件 {skills_dir}\FY4B-auto-download\scripts\skill_config.json，遍历 push_targets 数组，将每个推送目标的 channel/to/accountId 配置字段内嵌于 cron payload 中，以便之后的任务无需读取配置文件，减少一轮 LLM 响应以降低超时风险。

> 💡 **QQ 直接推送备用方案**：`fy4b_download.py` 运行完会**同步等待**调用 `fy4b_push_qq.py` 直接通过 QQ 开放平台 API 推送（不再异步 Popen，避免父进程退出后子进程被进程树清理导致推送丢失）。
> - **QClaw cron 执行时请加 `--skip-push` 参数**（`python fy4b_download.py --skip-push`），由 QClaw 自行推送，避免重复
> - Token 获取：`POST https://api.bot.qq.com/app/getAppAccessToken`，body: `{"appId":"YOUR_APP_ID","clientSecret":"YOUR_APP_SECRET}"}`（凭证已存 `skill_config.json` 的 `qqbot_credentials`）
> - 推送接口：`POST https://api.bot.qq.com/v2/users/{user_openid}/messages`，Header: `Authorization: QQBot <token>`，body: `{"msg_type":0,"content":"<消息内容>"}`
> - Token 有效期约 7200 秒；**401 时脚本自动调 `refresh_qqbot_access_token()` 刷新并重试**，无需手动干预
> - 报告来源：`fy4b_download.py` 将 Markdown 报告写入 `scripts/last_report.txt`，推送脚本优先读取该文件
> - 排查：推送脚本每次执行都会追加写入 `scripts/push_log.txt`，异步/终端环境下推送失败可查此日志

### 失败分析输出机制

脚本会根据执行路径输出失败分析：

1. **双批次路径（过期更新）**：`fy4b_direct_download.py` 的 `main()` 在汇总两批下载后，合并失败原因并输出 `[FAIL_ANALYSIS]` 标记行
2. **单批次路径（未过期）**：`main()` 在当前批次下载后，如果有失败则输出 `[FAIL_ANALYSIS]` 标记行

`[FAIL_ANALYSIS]` 标记后的内容由 `fy4b_download.py` 的 `parse_output()` 通过正则提取，并传入 `build_message()` 追加到推送消息末尾（Markdown 引用块格式）。

失败原因分类逻辑在 `fy4b_direct_download.py` 的 `_analyze_failures()` 中：
- HTTP 404 → "属上游资源问题，可忽略"
- HTTP 403 → "可能被限流或权限问题，稍后重试"
- 超时 → "网络波动导致，可重试"
- SSL → "证书问题，检查系统时间或网络环境"
- 连接失败 → "网络不通，请检查网络连接"
- 其他 → "需人工排查"

### 手动触发

用户说"测试 FY4B 下载"时，直接执行 `python {skill_dir}/scripts/fy4b_download.py`，解析日志，推送结果。

### 定时任务配置

首次配置时，创建 cron 任务：

- schedule: `30 9,15,21 * * *`（每天 09:30 / 15:30 / 21:30，北京时间）
- sessionTarget: `isolated`
- payload.kind: `agentTurn`
- payload.timeoutSeconds: `180`（原 120s 偏紧，LLM 首轮响应可达 90+ 秒）
- delivery.mode: `none`（推送由 agentTurn 内的 message 工具完成）

> ⚠️ 超时问题说明：Python 脚本本身只需 ~10 秒，但 LLM agent 首轮响应可达 90+ 秒。120s 超时下，LLM 第二轮响应（调用 message 推送）往往被 abort。已将超时提升至 180s 并精简 payload 以减少 LLM 处理时间。

参考 `references/cron-setup.md` 获取详细步骤。

## 推送消息格式

```markdown
📡 **FY4B 云图下载报告**

- ❌ 失败: **W**
- ✅ 成功: **X**
- ⏭️ 已存在: **Y**
- 🔮 未来图片: **Z**

⏰ 执行时间：{exec_time}

> 截止T：xxx.txt (N 个链接)
> N个文件HTTP XXX错误，xxx。
```

> 💡 当 `fail > 0` 时，`build_message()` 会自动在消息末尾追加失败分析（Markdown 引用块格式）。无失败时不追加。

## Skill 文件结构

```
FY4B-auto-download/
├── SKILL.md                          # 本文件
├── scripts/
│   ├── skill_config.json              # 统一配置文件（所有路径和参数）
│   ├── fy4b_download.py              # 入口脚本（cron 调用此文件）
│   ├── fy4b_direct_download.py       # 核心下载逻辑（被入口脚本调用）
│   ├── check_excel.py                # Excel/txt 状态检查模块
│   ├── create_download_links.py      # 链接生成 & Excel 更新模块
│   ├── fy4b_idm_import.py            # IDM 导入方案（备选）
│   └── templates/
│       └── url_template.md           # 下载链接 URL 模板
└── references/
    ├── cron-setup.md                 # 定时任务配置参考
    └── excel-structure.md            # Excel 结构参考
```

## 脚本依赖关系

```
fy4b_download.py          ← cron 定时调用
  └── fy4b_direct_download.py   ← subprocess 调用
        ├── check_excel.py           ← import
        └── create_download_links.py ← import
              └── templates/url_template.md ← 读取模板
```

### 失败分析数据流

```
download_links() 失败记录
  → _analyze_failures() 分类+建议（404/403/超时/SSL/连接失败/其他）
    → main() 输出 [FAIL_ANALYSIS] 标记行到 stdout
      → fy4b_download.py parse_output() 正则提取 [FAIL_ANALYSIS] 内容
        → build_message() 追加到消息末尾（Markdown 引用块格式）
          → 推送至 QQ/微信
```

所有 Python 脚本均在同一 `scripts/` 目录下，互相通过 `import` 引用，无需调整 `sys.path`。

## Excel 结构参考

- B1：当前批次过期时间（北京时间）
- B2→B3→...→B161：公式链（每行 +15 分钟）
- B162：下一批次过期时间（由脚本更新）
- 更新逻辑：B162 值 → 覆盖 B2 → Excel 内部公式链自动重算（B2→B3→...→B162）
- 停止条件：B1 ≥ 当前时间 - 12 小时（剩余有效期 ≥ 12h 即停止循环）

详见 `references/excel-structure.md`。

## 更新定时任务

当用户修改推送目标（如更换微信 ID）时：
1. 修改 `scripts/skill_config.json` 中的 `push_targets`
2. 执行 `cron update` 同步修改任务 payload（推送目标从配置读取，payload 无需变更）
3. 实际上 payload 指示 agent 读取配置文件，因此只需改配置文件即可，无需更新 cron

## 调试

- 脚本运行后会在 `scripts/skill_config.json` 中的 `{paths.urls_txt_file}/last_output.txt` 写入原始输出，方便排查解析失败问题
- 直接运行 `python fy4b_direct_download.py` 可查看完整日志