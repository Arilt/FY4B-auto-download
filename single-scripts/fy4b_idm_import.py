#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FY4B 云图 IDM 自动导入脚本

功能：
1. 检查 Excel/txt 是否存在
   - 都不存在 → 生成新 txt（根据模板）
   - txt 存在、Excel 不存在 → 保留 txt 当作当前批次
   - Excel 存在 → 正常流程
2. 检查待下载文件中的日期时间是否已过期
3. 调用 IDM 导入下载任务
4. 使用 Excel COM 接口更新文件并获取计算后的值

所有路径、软件位置、参数均从 fy4b_config.json 读取。
"""

import json
import re
import sys
import subprocess
from datetime import datetime, timedelta
from pathlib import Path

# 导入检查器和链接生成器
import check_excel
import create_download_links


# ===================== 加载配置 =====================
_CFG_PATH = Path(__file__).parent / "fy4b_config.json"
_SCRIPT_DIR = Path(__file__).parent

with open(_CFG_PATH, "r", encoding="utf-8") as _f:
    C = json.load(_f)

def _resolve_path(raw_path, base_dir=None):
    """智能路径解析：绝对路径直接使用，相对路径以 base_dir 为基准解析，~ 自动展开。"""
    if base_dir is None:
        base_dir = _SCRIPT_DIR
    # 接受 str 或 Path 对象
    if isinstance(raw_path, Path):
        p_raw = raw_path
    else:
        p_raw = Path(str(raw_path).strip())
    # 展开 ~
    p_raw = p_raw.expanduser()
    if p_raw.is_absolute():
        return p_raw.resolve()
    return (base_dir / p_raw).resolve()


# 路径（所有路径都走智能解析）
TXT_DIR    = _resolve_path(C["paths"]["urls_txt_file"])
EXCEL_FILE = _resolve_path(C["paths"]["Excel_file"])
IDM_PATH   = _resolve_path(C["paths"]["idm_program"])

# Excel 结构
_XL_SHEET     = C["Excel"]["sheet"]
_XL_B162      = C["Excel"]["B162单元格"]
_XL_B2        = C["Excel"]["B2单元格"]
_XL_B1        = C["Excel"]["B1单元格"]
_XL_LINK_COL  = C["Excel"]["链接列"]
_XL_ROW_START = C["Excel"]["链接起始行"]
_XL_ROW_END   = C["Excel"]["链接结束行"]

# txt 文件
TXT_PREFIX  = C["txt_regular"]["前缀"]
TXT_DATEFMT = C["txt_regular"]["日期格式"]


# ===================== 辅助函数 =====================
def log(msg):
    """带时间戳的日志（自动替换 emoji 避免编码问题）"""
    for emoji, tag in [
        ('❌', '[FAIL]'), ('✅', '[OK]'), ('📄', '[FILE]'),
        ('📅', '[DATE]'), ('🕐', '[TIME]'), ('📦', '[IDM]'),
        ('📊', '[EXCEL]'), ('📝', '[COPY]'), ('📋', '[LINKS]'),
        ('⏳', '[WAIT]'), ('📂', '[OPEN]'),
    ]:
        msg = msg.replace(emoji, tag)
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")


def ole_to_datetime(value):
    """Excel OLE 日期 / datetime → Python datetime（统一去除时区信息）"""
    if value is None:
        return None
    if isinstance(value, datetime):
        # 去除时区信息，避免与 offset-naive 的 datetime.now() 比较报错
        return value.replace(tzinfo=None)
    if isinstance(value, (int, float)):
        return datetime(1899, 12, 30) + timedelta(days=float(value))
    return None


def parse_txt_datetime(filename: str) -> datetime | None:
    """从截止T：YYYY年MM月DD日hhmm.txt 提取 datetime"""
    pat = rf"{re.escape(TXT_PREFIX)}(\d{{4}})年(\d{{2}})月(\d{{2}})日(\d{{2}})(\d{{2}})\.txt$"
    match = re.search(pat, filename)
    if match:
        y, mo, d, h, mi = match.groups()
        return datetime(int(y), int(mo), int(d), int(h), int(mi))
    return None


def read_txt_links(txt_path):
    """从 txt 读取链接列表"""
    with open(txt_path, "r", encoding="utf-8") as f:
        return [line.strip() for line in f if line.strip()]


# ===================== Excel 操作 =====================
def update_excel_and_file(txt_file):
    """使用 Excel COM 接口更新 Excel 和文本文件，返回 (成功?, 新文件名?)"""
    try:
        import win32com.client
    except ImportError:
        log("[FAIL] 需要安装 pywin32: pip install pywin32")
        return False, None

    excel = wb = None
    try:
        log("[EXCEL] 启动 Excel...")
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        log(f"[OPEN] 打开文件: {EXCEL_FILE}")
        wb = excel.Workbooks.Open(str(EXCEL_FILE))
        ws = wb.Worksheets(_XL_SHEET)

        # 循环: B162值粘贴→B2, 重算, 检查 B1 >= 当前时间
        now = datetime.now()
        iteration = 0
        max_iter = 50
        new_b1 = None

        cutoff = now - timedelta(hours=12)  # 停止条件：B1 >= now-12h

        while True:
            iteration += 1
            if iteration > max_iter:
                log(f"[WARN] 已循环 {max_iter} 次仍未满足条件，强制退出")
                break

            b162_value = ws.Range(_XL_B162).Value
            if b162_value is None:
                log(f"[FAIL] {_XL_B162} 为空，无法更新")
                return False, None

            # 值粘贴 B162 → B2
            ws.Range(_XL_B2).Value = b162_value
            log(f"[COPY] 第{iteration}次: {_XL_B162} → {_XL_B2}: {b162_value}")
            excel.Calculate()

            b1_value = ws.Range(_XL_B1).Value
            new_b1 = ole_to_datetime(b1_value) if b1_value else None
            log(f"[DATE] 第{iteration}次: B1 = {b1_value}")

            if new_b1 and new_b1 >= cutoff:
                log(f"[EXCEL] B1 >= {cutoff.strftime('%Y-%m-%d %H:%M')}，满足条件，停止循环")
                break
            else:
                log(f"[EXCEL] B1 < {cutoff.strftime('%Y-%m-%d %H:%M')}，继续循环...")

        # 收集链接
        e_values = []
        for row in range(_XL_ROW_START, _XL_ROW_END + 1):
            cell_value = ws.Range(f"{_XL_LINK_COL}{row}").Value
            if cell_value:
                e_values.append(str(cell_value))
        log(f"[LINKS] 收集到 {len(e_values)} 个链接")

        # 读取最终的 B1
        b1_value = ws.Range(_XL_B1).Value
        log(f"[DATE] 最终 B1 值: {b1_value} (循环{iteration}次)")

        wb.Save()
        log("[OK] Excel 已保存")

        # 更新文本文件内容
        with open(txt_file, "w", encoding="utf-8") as f:
            f.write("\n".join(e_values))
        log(f"[OK] 文本文件已更新: {txt_file}")

        # 更新文件名
        new_filename = None
        if b1_value:
            new_dt = ole_to_datetime(b1_value)
            if new_dt:
                new_filename = f"{TXT_PREFIX}{new_dt.strftime(TXT_DATEFMT)}.txt"
                new_txt_path = TXT_DIR / new_filename
                if txt_file != new_txt_path:
                    txt_file.rename(new_txt_path)
                    log(f"[OK] 文件已重命名: {new_filename}")

        return True, new_filename

    except Exception as e:
        log(f"[FAIL] 更新 Excel/文本文件失败: {e}")
        import traceback
        traceback.print_exc()
        return False, None
    finally:
        if wb:
            wb.Close(SaveChanges=False)
        if excel:
            excel.Quit()


# ===================== 主流程 =====================
def check_and_import():
    """主流程"""
    log("=" * 50)
    log("FY4B IDM 自动导入任务开始")
    log(f"[CFG] 配置文件: {_CFG_PATH}")
    log("=" * 50)

    # 1. 检查 Excel/txt 是否存在
    excel_ok, txt_ok, txt_path, txt_dt = check_excel.check_files_exist()
    log(f"[CHECK] Excel: {'存在' if excel_ok else '不存在'}")
    log(f"[CHECK] txt: {'存在' if txt_ok else '不存在'}")

    now = datetime.now()

    # 2. 根据文件存在情况决定行为
    just_generated = False  # 标记是否刚生成了新 txt

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
    
    # 3. 检查是否过期
    if txt_path is None:
        log("[FAIL] 无法获取 txt 文件")
        return False
    
    log(f"[FILE] 文件: {txt_path.name}")
    
    if txt_dt:
        log(f"[DATE] 过期时间: {txt_dt.strftime('%Y-%m-%d %H:%M')}")
    else:
        log(f"[FAIL] 无法从文件名提取日期时间: {txt_path.name}")
        return False
    
    log(f"[TIME] 当前时间: {now.strftime('%Y-%m-%d %H:%M')}")

    # 过期时间 + 1.5 小时缓冲
    expire_with_buffer = txt_dt + timedelta(hours=-41.5-11/30)

    if just_generated:
        # 刚生成的 txt，直接导入（不过期检查）
        log("[INFO] 刚生成的新 txt，直接导入")
    # elif now <= txt_dt:
    elif now <= expire_with_buffer:
        log("[WAIT] 当前时间未超过文件时间，无需导入")
        return True
    else:
        log("[OK] 当前时间已超过文件时间，开始导入流程")

    # 4. 调用 IDM 导入
    log(f"[IDM] 调用 IDM 导入: {txt_path}")
    try:
        result = subprocess.run(
            [str(IDM_PATH), "/s", "/import", str(txt_path)],
            capture_output=True,
            text=True,
            timeout=60
        )
        log(f"[OK] IDM 导入命令已执行 (返回码: {result.returncode})")
    except subprocess.TimeoutExpired:
        log("[FAIL] IDM 导入超时")
        return False
    except Exception as e:
        log(f"[FAIL] IDM 导入失败: {e}")
        return False

    # 5. 更新 Excel 或生成新 txt
    if excel_ok:
        # Excel 存在 → 用 Excel 更新
        log("[EXCEL] 开始更新 Excel 和文本文件...")
        success, new_filename = update_excel_and_file(txt_path)
        if not success:
            return False
    else:
        # Excel 不存在 → 用模板生成新 txt
        log("[GEN] Excel 不存在，用模板生成新 txt")
        try:
            new_txt_path = create_download_links.generate_and_save_txt(start_time=txt_dt)
            log(f"[OK] 生成文件: {new_txt_path.name}")
            # 删除旧 txt
            if txt_path and txt_path.exists():
                txt_path.unlink()
                log(f"[DELETE] 旧文件: {txt_path.name}")
        except Exception as e:
            log(f"[FAIL] 生成 txt 失败: {e}")
            return False

    log("=" * 50)
    log("[OK] FY4B IDM 自动导入任务完成")
    log("=" * 50)
    return True


if __name__ == "__main__":
    sys.exit(0 if check_and_import() else 1)