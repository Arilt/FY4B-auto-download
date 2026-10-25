#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FY4B 云图下载链接生成 & Excel 更新模块

提供两个主入口：
  - generate_and_save_txt(start_time=None) → 生成链接并保存到 txt
  - update_excel_and_generate_txt(excel_file=None) → 循环更新 Excel 达标后生成新 txt

所有参数从 fy4b_config.json 读取。
"""

import json
import re
import shutil
from datetime import datetime, timedelta
from pathlib import Path


# ===================== 加载配置 =====================
_SCRIPT_DIR = Path(__file__).parent
_CFG_PATH = _SCRIPT_DIR / "fy4b_config.json"

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


# ===================== 配置常量 =====================
# 路径
TXT_DIR      = _resolve_path(C["paths"]["urls_txt_file"])
TEMPLATE_FILE = _resolve_path(C["paths"]["url_template_file"])

# txt 文件
TXT_PREFIX            = C["txt_regular"]["前缀"]
TXT_DATEFMT           = C["txt_regular"]["日期格式"]
LINK_COUNT            = C["txt_regular"]["链接数量"]
OFFSET_HOURS          = C["txt_regular"]["时间偏移小时"]
TOTAL_HOURS           = C["txt_regular"]["总时长小时"]
INTERVAL_MINUTES      = C["txt_regular"]["链接间隔分钟"]
EXPIRY_OFFSET_MINUTES = C["txt_regular"]["截止时间偏移分钟"]
TIMEZONE_OFFSET_HOURS = C["txt_regular"]["时区偏移小时"]

# Excel 结构
EXCEL_FILE   = _resolve_path(C["paths"]["Excel_file"])
_XL_SHEET    = C["Excel"]["sheet"]
_XL_B162     = C["Excel"]["B162单元格"]
_XL_B2       = C["Excel"]["B2单元格"]
_XL_B1       = C["Excel"]["B1单元格"]
_XL_LINK_COL = C["Excel"]["链接列"]
_XL_ROW_START = C["Excel"]["链接起始行"]
_XL_ROW_END   = C["Excel"]["链接结束行"]


# ===================== Excel 操作 =====================
def ole_to_datetime(value):
    """Excel OLE 日期 / datetime → Python datetime（统一去除时区信息）。"""
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.replace(tzinfo=None)
    if isinstance(value, (int, float)):
        return datetime(1899, 12, 30) + timedelta(days=float(value))
    return None


def format_txt_name(dt):
    """datetime → 截止T：2026年04月28日0907.txt"""
    return f"{TXT_PREFIX}{dt.strftime(TXT_DATEFMT)}.txt"


def update_excel_and_generate_txt(excel_file=None, cutoff_hours=12):
    """
    循环更新 Excel（B162值粘贴→B2，重算）直到 B1 >= 当前时间-cutoff_hours，
    然后将 E 列链接写入新的截止T.txt。返回新 txt 路径，或 None 失败。

    Args:
        excel_file: Excel 文件路径，默认使用配置中的路径
        cutoff_hours: 停止条件，B1 >= 当前时间 - 此值（小时），默认 12
    """
    import win32com.client

    if excel_file is None:
        excel_file = EXCEL_FILE

    def log(msg):
        print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")

    log("[EXCEL] 启动更新流程...")

    excel = wb = ws = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        wb = excel.Workbooks.Open(str(excel_file))
        ws = wb.Worksheets(_XL_SHEET)
        excel.Calculate()

        now = datetime.now()
        cutoff = now - timedelta(hours=cutoff_hours)
        iteration = 0
        max_iter = 50
        new_b1 = None

        while True:
            iteration += 1
            if iteration > max_iter:
                log(f"[WARN] 已循环 {max_iter} 次仍未满足条件，强制退出")
                break

            b162_val = ws.Range(_XL_B162).Value
            if b162_val is None:
                log(f"[FAIL] {_XL_B162} 为空，无法更新")
                return None

            ws.Range(_XL_B2).Value = b162_val
            excel.Calculate()

            new_b1 = ole_to_datetime(ws.Range(_XL_B1).Value)
            if not new_b1:
                log("[FAIL] B1 值无法解析")
                return None

            log(f"[EXCEL] 第{iteration}次: B162→B2, B1={new_b1.strftime('%Y-%m-%d %H:%M')}")

            if new_b1 >= cutoff:
                log(f"[EXCEL] B1 >= {cutoff.strftime('%Y-%m-%d %H:%M')}，满足条件，停止循环")
                break
            else:
                log(f"[EXCEL] B1 < {cutoff.strftime('%Y-%m-%d %H:%M')}，继续循环...")

        # 收集链接
        links = []
        for row in range(_XL_ROW_START, _XL_ROW_END + 1):
            v = ws.Range(f"{_XL_LINK_COL}{row}").Value
            if v:
                links.append(str(v))
        log(f"[LINKS] 收集到 {len(links)} 个链接")

        if not new_b1 or not links:
            log("[FAIL] 更新后 B1 或链接为空")
            return None

        wb.Save()
        log(f"[EXCEL] 更新完成，新过期时间: {new_b1.strftime('%Y-%m-%d %H:%M')} (循环{iteration}次)")

        new_txt_path = TXT_DIR / format_txt_name(new_b1)
        TXT_DIR.mkdir(parents=True, exist_ok=True)
        with open(new_txt_path, "w", encoding="utf-8") as f:
            f.write("\n".join(links))
        log(f"[OK] 新文件: {new_txt_path.name} ({len(links)} 个链接)")

        return new_txt_path

    except Exception as e:
        log(f"[FAIL] Excel 更新失败: {e}")
        import traceback
        traceback.print_exc()
        return None
    finally:
        if wb:
            wb.Close(SaveChanges=False)
        if excel:
            excel.Quit()


# ===================== 模板解析 =====================
def parse_template() -> str:
    """从 url_template.md 解析链接模板。"""
    if not TEMPLATE_FILE.exists():
        raise FileNotFoundError(f"模板文件不存在: {TEMPLATE_FILE}")
    with open(TEMPLATE_FILE, "r", encoding="utf-8") as f:
        content = f.read()
    code_block_match = re.search(r"```\s*\n?(https?://[^\s`]+)\n?```", content, re.DOTALL)
    if code_block_match:
        return code_block_match.group(1).strip()
    link_match = re.search(r"https?://[^\s`]+", content)
    if link_match:
        return link_match.group(0).strip()
    raise ValueError(f"模板文件中未找到有效链接: {TEMPLATE_FILE}")


def format_link(template: str, dt: datetime) -> str:
    """将模板中的占位符替换为实际时间值。"""
    result = template
    result = result.replace("YYYY", dt.strftime("%Y"))
    result = result.replace("MM",   dt.strftime("%m"))
    result = result.replace("DD",     dt.strftime("%d"))
    result = result.replace("hh",     dt.strftime("%H"))
    result = result.replace("mm",     dt.strftime("%M"))
    return result


# ===================== 链接生成 =====================
def generate_links(start_time: datetime = None) -> list[str]:
    """生成下载链接列表（默认从当前时间-25h整点开始-时区偏移）。"""
    template = parse_template()
    if start_time is None:
        start_time = (datetime.now() - timedelta(hours=OFFSET_HOURS + TIMEZONE_OFFSET_HOURS)).replace(
            minute=0, second=0, microsecond=0
        )
    return [format_link(template, start_time + timedelta(minutes=i * INTERVAL_MINUTES))
            for i in range(LINK_COUNT)]


def generate_txt_filename(links: list[str]) -> str:
    """根据最后一个链接的时间生成 txt 文件名（最后链接时间 + 22分钟 + 时区偏移（东八区）= 截止时间）。"""
    last_link = links[-1]
    match = re.search(r"(\d{4})(\d{2})(\d{2})(\d{2})(\d{2})00000\.JPG$", last_link, re.IGNORECASE)
    if not match:
        raise ValueError(f"无法从链接中提取时间: {last_link}")
    y, mo, d, h, mi = match.groups()
    last_dt = datetime(int(y), int(mo), int(d), int(h), int(mi))
    expiry_dt = last_dt + timedelta(minutes=(EXPIRY_OFFSET_MINUTES + TIMEZONE_OFFSET_HOURS*60))  #  时区偏移
    return f"{TXT_PREFIX}{expiry_dt.strftime(TXT_DATEFMT)}.txt"


def generate_and_save_txt(start_time: datetime = None) -> Path:
    """生成链接并保存到 txt 文件。"""
    if start_time is not None:
        start_time = start_time - timedelta(minutes=EXPIRY_OFFSET_MINUTES - INTERVAL_MINUTES + TIMEZONE_OFFSET_HOURS*60)
    links = generate_links(start_time)
    filename = generate_txt_filename(links)
    txt_path = TXT_DIR / filename
    TXT_DIR.mkdir(parents=True, exist_ok=True)
    with open(txt_path, "w", encoding="utf-8") as f:
        f.write("\n".join(links))
    return txt_path


# ===================== 主入口 =====================
if __name__ == "__main__":
    import sys
    print("=" * 50)
    print("FY4B 云图链接生成器")
    print(f"[CFG] 配置文件: {_CFG_PATH}")
    print("=" * 50)
    try:
        template = parse_template()
        print(f"[TEMPLATE] {template}")
        txt_path = generate_and_save_txt()
        print(f"[OK] 生成文件: {txt_path.name}")
        print(f"[INFO] 包含 {LINK_COUNT} 个链接")
        links = generate_links()
        print(f"[TIME] 时间范围: {links[0][-20:-4]} → {links[-1][-20:-4]}")
        sys.exit(0)
    except Exception as e:
        print(f"[FAIL] {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)