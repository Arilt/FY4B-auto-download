#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FY4B 云图自动下载脚本

流程:
1. 检查 Excel/txt 是否存在
   - 都不存在 → 生成新 txt（根据模板）
   - txt 存在、Excel 不存在 → 保留 txt 当作当前批次
   - Excel 存在 → 正常流程
2. 查找截止T：{B1}.txt 文件，提取过期时间
3. 当前时间 > B1（已过期）:
   a. 读取旧 txt → 下载到 {临时目录}
   b. 更新 Excel (B162→B2) 或 生成新 txt
   c. 将 {临时目录} 所有图片移动至 {目标目录}
   d. 移动完成后，生成新的截止T txt
   e. 预下载下一批次到 {临时目录}
4. 当前时间 <= B1（未过期）:
   a. 读取现有 txt → 下载到 {临时目录}

所有路径、软件位置、参数均从 skill_config.json 读取，改配置即可，无需动脚本。
"""

import json
import re
import sys
import time
import shutil
from pathlib import Path
from datetime import datetime, timedelta

# 导入检查器和链接生成器
import check_excel
import create_download_links


# ===================== 加载配置 =====================
_CFG_PATH = Path(__file__).parent / "skill_config.json"
_SCRIPT_DIR = Path(__file__).parent

with open(_CFG_PATH, "r", encoding="utf-8") as _f:
    C = json.load(_f)


def _resolve_path(raw_path, base_dir=None):
    """智能路径解析：绝对路径直接使用，相对路径以 base_dir 为基准解析，~ 自动展开。"""
    if base_dir is None:
        base_dir = _SCRIPT_DIR
    if isinstance(raw_path, Path):
        p_raw = raw_path
    else:
        p_raw = Path(str(raw_path).strip())
    p_raw = p_raw.expanduser()
    if p_raw.is_absolute():
        return p_raw.resolve()
    return (base_dir / p_raw).resolve()


# 路径（所有路径都走智能解析）
STAGING_DIR = _resolve_path(C["paths"]["staging_dir"])
TARGET_DIR  = _resolve_path(C["paths"]["target_dir"])
TXT_DIR     = _resolve_path(C["paths"]["urls_txt_file"])
EXCEL_FILE  = _resolve_path(C["paths"]["Excel_file"])
IDM_PATH    = _resolve_path(C["paths"]["idm_program"])

# 下载参数
_DL_TIMEOUT   = C["download_setting"]["超时秒数"]
_DL_PQL_EVERY = C["download_setting"]["每N个请求暂停"]
_DL_PQL_SEC   = C["download_setting"]["暂停秒数"]
_DL_PROGRESS  = C["download_setting"]["进度汇报间隔"]
_DL_UA        = C["download_setting"]["UserAgent"]
_DL_EXTS      = set(C["download_setting"]["图片后缀"])


# ===================== 辅助函数 =====================
def log(msg):
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")


def parse_txt_datetime(filename):
    """截止T：YYYY年MM月DD日HHMM.txt → datetime"""
    pat = rf"{re.escape(create_download_links.TXT_PREFIX)}(\d{{4}})年(\d{{2}})月(\d{{2}})日(\d{{2}})(\d{{2}})\.txt$"
    m = re.search(pat, filename)
    if m:
        y, mo, d, h, mi = m.groups()
        return datetime(int(y), int(mo), int(d), int(h), int(mi))
    return None


def read_txt_links(txt_path):
    """从 txt 读取链接列表"""
    with open(txt_path, "r", encoding="utf-8") as f:
        return [line.strip() for line in f if line.strip()]


def is_future_link(url: str, future_minutes: int = -47) -> bool:
    """
    检查链接是否指向"未来"图片（文件名时间 - 8h > 当前时间 + future_minutes）。
    文件名时间为 UTC+0，需转换为北京时间（UTC+8）后比较。
    未来图片尚未生成，下载会得到 404 或无效内容，直接跳过。
    """
    import re as _re
    filename = url.split("/")[-1]
    m = _re.search(r"(\d{4})(\d{2})(\d{2})(\d{2})(\d{2})00000\.JPG$", filename, _re.IGNORECASE)
    if not m:
        return False
    y, mo, d, h, mi = m.groups()
    link_dt_utc = datetime(int(y), int(mo), int(d), int(h), int(mi))
    # 文件名是 UTC+0，转换为北京时间（UTC+8）
    link_dt_beijing = link_dt_utc + timedelta(hours=8)
    return link_dt_beijing > datetime.now() + timedelta(minutes=future_minutes)


def move_staging_to_target():
    """将 STAGING_DIR 中所有图片移动到 TARGET_DIR"""
    if not STAGING_DIR.exists():
        return 0
    moved = skipped = 0
    for f in STAGING_DIR.iterdir():
        if f.suffix.lower() not in _DL_EXTS:
            continue
        TARGET_DIR.mkdir(parents=True, exist_ok=True)
        dest = TARGET_DIR / f.name
        try:
            if not f.exists():
                log(f"[MOVE SKIP] 源文件不存在（可能被其他进程移动）: {f.name}")
                skipped += 1
                continue
            shutil.move(str(f), str(dest))
            moved += 1
        except Exception as e:
            if "cannot find" in str(e).lower() or "not found" in str(e).lower():
                log(f"[MOVE SKIP] 文件不存在: {f.name}")
                skipped += 1
            else:
                log(f"[MOVE FAIL] {f.name}: {e}")
    if moved or skipped:
        log(f"[MOVE] 移动 {moved} 个，跳过 {skipped} 个 → {TARGET_DIR}")
    return moved


# ===================== 下载 =====================
def download_links(links, batch_label=""):
    """批量下载到 STAGING_DIR，跳过已存在于 STAGING_DIR 或 TARGET_DIR 的文件及未来图片。

    返回 (fail, success, skip_exist, skip_future, fail_details) 五元组。
    fail_details 为失败原因分析字符串（空字符串表示无失败）。
    """
    fail = success = skip_exist = skip_future = 0
    fail_records = []  # 收集失败记录用于原因分析
    total = len(links)

    log(f"[DOWNLOAD] 共 {total} 个文件，目标: {STAGING_DIR}")
    STAGING_DIR.mkdir(parents=True, exist_ok=True)

    for i, url in enumerate(links, 1):
        filename = url.split("/")[-1]
        # 跳过未来图片（文件尚未生成，下载会得到 404 或无效内容）
        if is_future_link(url):
            skip_future += 1
            continue

        staging_path = STAGING_DIR / filename
        target_path = TARGET_DIR / filename

        if staging_path.exists() or target_path.exists():
            skip_exist += 1
            continue

        try:
            import urllib.request
            req = urllib.request.Request(url, headers={"User-Agent": _DL_UA})
            with urllib.request.urlopen(req, timeout=_DL_TIMEOUT) as resp:
                with open(staging_path, "wb") as f:
                    f.write(resp.read())
            success += 1
            if success % _DL_PROGRESS == 0 or i == total:
                log(f"[PROGRESS] {i}/{total} | 失败: {fail} | 成功: {success} | 已存在: {skip_exist} | 未来: {skip_future}")
        except Exception as e:
            fail += 1
            err_str = str(e)
            log(f"[FAIL] {filename}: {err_str}")
            fail_records.append((filename, err_str))

        if i % _DL_PQL_EVERY == 0:
            time.sleep(_DL_PQL_SEC)

    # 失败原因分析
    fail_details = _analyze_failures(fail_records, batch_label)
    # if fail_details:
    #     log(fail_details)

    log(f"[DONE] 失败: {fail} | 成功: {success} | 已存在: {skip_exist} | 未来: {skip_future}")
    return fail, success, skip_exist, skip_future, fail_details


def _analyze_failures(fail_records, batch_label=""):
    """分析失败记录，返回人类可读的原因汇总字符串。

    按错误类型分类统计，输出类似：
        截止T: xxx.txt (160 个链接)
        1个文件HTTP 404错误，属上游资源问题，可忽略。
    """
    if not fail_records:
        return ""

    # 按错误类型归类
    from collections import Counter
    error_types = Counter()
    for _filename, err_str in fail_records:
        err_lower = err_str.lower()
        if "404" in err_lower or "not found" in err_lower:
            error_types["HTTP 404"] += 1
        elif "403" in err_lower or "forbidden" in err_lower:
            error_types["HTTP 403"] += 1
        elif "timeout" in err_lower or "timed out" in err_lower:
            error_types["超时"] += 1
        elif "ssl" in err_lower or "certificate" in err_lower:
            error_types["SSL"] += 1
        elif "connection" in err_lower or "refused" in err_lower:
            error_types["连接失败"] += 1
        else:
            error_types["其他错误"] += 1

    parts = []
    if batch_label:
        parts.append(f"{batch_label} ({len(fail_records)} 个链接)")
    for err_type, count in error_types.items():
        suggestion = _failure_suggestion(err_type)
        parts.append(f"{count}个文件{err_type}错误，{suggestion}。")
    return "\n".join(parts)


def _failure_suggestion(err_type):
    """根据错误类型给出建议"""
    suggestions = {
        "HTTP 404": "属上游资源问题，可忽略",
        "HTTP 403": "可能被限流或权限问题，稍后重试",
        "超时": "网络波动导致，可重试",
        "SSL": "证书问题，检查系统时间或网络环境",
        "连接失败": "网络不通，请检查网络连接",
        "其他错误": "需人工排查",
    }
    return suggestions.get(err_type, "需人工排查")


# ===================== 主流程 =====================
def main():
    log("=" * 50)
    log("FY4B 云图自动下载")
    log(f"[CFG] 配置文件: {_CFG_PATH}")
    log("=" * 50)

    # 1. 检查 Excel/txt 是否存在
    excel_ok, txt_ok, txt_path, txt_dt = check_excel.check_files_exist()
    log(f"[CHECK] Excel: {'存在' if excel_ok else '不存在'}")
    log(f"[CHECK] txt: {'存在' if txt_ok else '不存在'}")

    now = datetime.now()
    need_update = False
    just_generated = False  # 标记是否刚生成了新 txt

    # 2. 根据文件存在情况决定行为
    if not excel_ok and not txt_ok:
        # 场景3：都不存在 → 生成新 txt
        log("[GEN] Excel 和 txt 均不存在，根据模板生成新 txt")
        try:
            txt_path = create_download_links.generate_and_save_txt()
            txt_dt = parse_txt_datetime(txt_path.name)
            just_generated = True
            log(f"[OK] 生成文件: {txt_path.name}")
        except Exception as e:
            log(f"[FAIL] 生成 txt 失败: {e}")
            return False

    elif not excel_ok and txt_ok:
        # 场景4：txt 存在、Excel 不存在 → 保留 txt 当作当前批次
        log(f"[KEEP] txt 存在但 Excel 不存在，保留当前 txt 作为当前批次")
        log(f"[INFO] 文件: {txt_path.name}")

    # 3. 判断是否需要更新（刚生成的 txt 不需要再更新）
    if just_generated:
        log("[INFO] 刚生成的新 txt，直接下载当前批次")
    elif txt_dt:
        # 过期时间 + 1.2 小时缓冲
        expire_with_buffer = txt_dt + timedelta(hours=1.2)
        log(f"[CHECK] 过期时间(含1.2h缓冲): {expire_with_buffer.strftime('%Y-%m-%d %H:%M')}")
        log(f"[CHECK] 过期时间: {txt_dt.strftime('%Y-%m-%d %H:%M')}")
        log(f"[CHECK] 当前时间: {now.strftime('%Y-%m-%d %H:%M')}")
        # if now > txt_dt:
        if now > expire_with_buffer:
            need_update = True
            log("[INFO] 已过期")
        else:
            log("[INFO] 未过期，直接下载")
    else:
        need_update = True
        log("[CHECK] 未找到截止T.txt，需更新")

    # 4. 读取 txt 并下载到临时目录
    if txt_path and txt_path.exists():
        links = read_txt_links(txt_path)
        log(f"[DOWNLOAD] 当前批次 → {txt_path.name} ({len(links)} 个链接)")
    elif need_update:
        links = None
    else:
        log("[FAIL] 无截止T.txt 且无需更新，退出")
        return False

    if links is not None:
        fail, ok, skip_exist, skip_future, fail_details = download_links(links, batch_label=txt_path.name if txt_path else "")
        log(f"[DONE] 当前批次 | 失败: {fail} | 成功: {ok} | 已存在: {skip_exist} | 未来: {skip_future}")
        # 单批次路径也输出失败分析标记
        if fail_details:
            log(f"[FAIL_ANALYSIS]{fail_details}")

    # 5. 过期 → 更新 Excel/生成新 txt → 移动 → 删旧txt → 预下载下一批次
    if need_update:
        old_txt_path = txt_path

        if excel_ok:
            # Excel 存在 → 调用 create_download_links 更新 Excel 并生成新 txt
            new_txt_path = create_download_links.update_excel_and_generate_txt(
                excel_file=EXCEL_FILE,
                cutoff_hours=12
            )
        else:
            # Excel 不存在 → 用模板生成新 txt
            log("[GEN] Excel 不存在，用模板生成新 txt")
            try:
                new_txt_path = create_download_links.generate_and_save_txt(start_time=txt_dt)
                log(f"[OK] 生成文件: {new_txt_path.name}")
            except Exception as e:
                log(f"[FAIL] 生成 txt 失败: {e}")
                return False

        if not new_txt_path:
            log("[FAIL] 更新失败，退出")
            return False

        moved = move_staging_to_target()
        log(f"[MOVE] 本次移动 {moved} 个文件")

        if old_txt_path and old_txt_path.exists():
            old_txt_path.unlink()
            log(f"[DELETE] 旧文件: {old_txt_path.name}")

        new_links = read_txt_links(new_txt_path)
        log(f"[DOWNLOAD] 下一批次 → {new_txt_path.name} ({len(new_links)} 个链接)")
        fail2, ok2, skip_exist2, skip_future2, fail_details2 = download_links(new_links, batch_label=new_txt_path.name)
        log(f"[DONE] 下一批次预下载 | 失败: {fail2} | 成功: {ok2} | 已存在: {skip_exist2} | 未来: {skip_future2}")

        # # 汇总两批下载的总数，供入口脚本正则捕获
        # total_new = ok + ok2
        # total_skip_exist = skip_exist + skip_exist2
        # total_skip_future = skip_future + skip_future2
        # total_fail = fail + fail2
        # log(f"[DONE] 汇总 | 成功: {total_new} | 已存在: {total_skip_exist} | 未来: {total_skip_future} | 失败: {total_fail}")
        # 合并失败原因（用特殊标记 [FAIL_ANALYSIS] 供外层解析）
        all_fail_details = "\n".join(filter(None, [fail_details, fail_details2]))
        if all_fail_details:
            log(f"[FAIL_ANALYSIS]{all_fail_details}")

    log("=" * 50)
    log("[OK] 任务完成")
    log("=" * 50)
    return True


if __name__ == "__main__":
    sys.exit(0 if main() else 1)