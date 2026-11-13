#!/usr/bin/env python3
"""
微信批量发送助手 — Windows 版（uiautomation 方案）
依赖: pip install uiautomation pywin32 Pillow openpyxl
核心：uiautomation 直接绑定微信 4.0+ 固定 UIA 控件结构，
      剪贴板粘贴输入文字（规避特殊字符/风控问题），
      双重策略打开会话（会话列表直接点 + 搜索框降级）。
用法: python scripts/wechat_send_win.py [--dry]
"""
from __future__ import annotations

import argparse
import json
import os
import random
import subprocess
import sys
import time
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import Optional


def _real_home() -> Path:
    """返回真实用户 home 目录"""
    if sys.platform == "win32":
        return Path(os.environ.get("USERPROFILE",
                                   os.environ.get("HOMEDRIVE", "C:\\") + "\\Users\\" + os.environ.get("USERNAME", "")))
    import pwd
    return Path(pwd.getpwuid(os.getuid()).pw_dir)


# ─── 依赖安装 ──────────────────────────────────────────
def _install_and_import(module, package=None):
    """尝试导入，失败则自动安装"""
    try:
        return __import__(module)
    except ImportError:
        pypi_name = package or module
        print(f"[信息] 正在安装 {pypi_name}...")
        r = subprocess.run(
            [sys.executable, "-m", "pip", "install", pypi_name, "--quiet"],
            capture_output=True, text=True
        )
        if r.returncode != 0:
            print(f"[ERROR] pip install {pypi_name} 失败: {r.stderr}", file=sys.stderr)
            sys.exit(1)
        return __import__(module)


# ─── UTF-8 编码保障 ────────────────────────────────────
try:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


# ─── 全局开关 ───────────────────────────────────────────
VERBOSE = False


def _log(msg: str):
    if VERBOSE:
        print(msg)


def _rsleep(lo: float = 0.2, hi: float = 0.5):
    """随机延时（降低风控风险）"""
    time.sleep(random.uniform(lo, hi))


# ─── uiautomation 导入 ─────────────────────────────────
# Windows 专用，macOS 不会运行到这里
try:
    import uiautomation as auto
except ImportError:
    _install_and_import("uiautomation")
    import uiautomation as auto

try:
    import win32clipboard
    import win32con
    from PIL import Image
    from io import BytesIO
    _HAS_WIN32 = True
except ImportError:
    _HAS_WIN32 = False

if _HAS_WIN32:
    from wechat_visual_win import VisualChat, find_window

# ─── 微信 UIA 常量（微信 4.0+ 固定结构）──────────────────
WX_CLASS   = "mmui::MainWindow"   # 微信主窗口 ClassName
WX_TITLE   = "微信"               # 微信主窗口 Name
WX_SESSION_CLASS = "mmui::ChatSessionCell"  # 会话列表单项 ClassName
# 置顶会话 Name 末尾会有此后缀（C# SDK 发现的规律）
WX_TOP_SUFFIX = "置顶"

# ─── 核心 UI 操作 ──────────────────────────────────────

def _get_wx() -> "auto.WindowControl":
    """
    获取微信主窗口（uiautomation），不到 3 行核心逻辑。
    若窗口不可见，尝试 Ctrl+Alt+W 唤醒。
    """
    wx = auto.WindowControl(searchDepth=1, Name=WX_TITLE, ClassName=WX_CLASS)
    if wx.Exists(0, 0):
        _log(f"[INFO] 找到微信主窗口 (ClassName={WX_CLASS})")
        return wx

    # 尝试 Ctrl+Alt+W 唤醒托盘微信
    _log("[INFO] 微信窗口不可见，尝试 Ctrl+Alt+W 唤醒")
    auto.SendKeys("{Ctrl}{Alt}w", waitTime=0.1)
    time.sleep(1.0)

    wx = auto.WindowControl(searchDepth=1, Name=WX_TITLE, ClassName=WX_CLASS)
    if wx.Exists(0, 0):
        _log("[INFO] 成功唤醒微信窗口")
        return wx

    raise RuntimeError("未找到微信窗口，请确保微信已启动并登录")


def _set_clipboard_text(text: str, retries: int = 3) -> bool:
    """安全写入剪贴板文字（带重试 + 校验）"""
    for attempt in range(retries):
        try:
            auto.SetClipboardText(text)
            time.sleep(0.05)
            got = auto.GetClipboardText()
            if got == text:
                return True
            _log(f"[WARN] 剪贴板校验失败（第 {attempt+1} 次），重试")
        except Exception as e:
            _log(f"[WARN] 剪贴板写入异常: {e}，重试")
        time.sleep(0.1)
    _log("[ERROR] 剪贴板写入多次失败")
    return False


def _copy_image_to_clipboard(image_path: str, retries: int = 3) -> bool:
    """将图片写入剪贴板（BMP 格式，微信粘贴需要）"""
    if not _HAS_WIN32:
        _log("[WARN] 缺少 pywin32/Pillow，图片发送不可用")
        return False
    for attempt in range(retries):
        clipboard_opened = False
        try:
            img = Image.open(image_path).convert("RGB")
            buf = BytesIO()
            img.save(buf, "BMP")
            data = buf.getvalue()[14:]   # 去掉 BMP 文件头
            buf.close()
            win32clipboard.OpenClipboard()
            clipboard_opened = True
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
            win32clipboard.CloseClipboard()
            clipboard_opened = False
            _log("[INFO] 图片已写入剪贴板")
            return True
        except Exception as e:
            if clipboard_opened:
                try:
                    win32clipboard.CloseClipboard()
                except Exception:
                    pass
            _log(f"[WARN] 图片剪贴板写入失败 (第{attempt+1}次): {e}")
            time.sleep(0.2)
    return False


def _open_chat(wx: "auto.WindowControl", target: str) -> bool:
    """
    打开目标会话（双重策略）：
    策略1：从会话列表直接点击（快）— 适合已有聊天记录的联系人
    策略2：搜索框输入（降级）— 适合新联系人或列表中找不到的

    关键细节（来自 WeChatAuto.SDK 源码）：
    - 会话列表 ListItem 的 Name 可能带"置顶"后缀，匹配时需处理
    - 实际可点击的是 ListItem 子控件 /Pane/Button，不是 ListItem 本身
    - LAVARONG 方案：AutomationId = "session_item_{name}"（更快）
    """
    wx.SetActive()
    _rsleep(0.2, 0.4)

    # ── 策略1：AutomationId 直接点击（微信 4.0 固定格式）──
    aid = f"session_item_{target}"
    try:
        item = wx.Control(ClassName=WX_SESSION_CLASS, AutomationId=aid, searchDepth=15)
        if item.Exists(0, 0):
            # 检查是否已选中（SelectionItemPattern，id=10010）
            try:
                pattern = item.GetPattern(10010)
                if pattern and getattr(pattern, "IsSelected", False):
                    _log(f"[INFO] 会话 '{target}' 已选中，跳过点击")
                    return True
            except Exception:
                pass
            _log(f"[INFO] 策略1：从会话列表直接点击 '{target}'")
            item.Click()
            _rsleep(0.3, 0.5)
            return True
    except Exception as e:
        _log(f"[DEBUG] 策略1 失败: {e}")

    # ── 策略2：搜索框输入（降级）──
    _log(f"[INFO] 策略2：使用搜索框查找 '{target}'")

    # 先切换到聊天主界面 Tab（确保搜索框可用）
    try:
        nav = wx.FindFirstByXPath("//Button[@Name='聊天']")
        if nav and nav.Exists(0, 0):
            nav.Click()
            _rsleep(0.2, 0.3)
    except Exception:
        pass

    # 找搜索框
    search_box = wx.EditControl(Name="搜索")
    if not search_box.Exists(0, 0):
        raise RuntimeError("未找到微信搜索框（请确认微信已切换到聊天界面）")

    search_box.Click()
    _rsleep(0.2, 0.3)

    # 剪贴板粘贴输入（避免特殊字符问题）
    if _set_clipboard_text(target):
        search_box.SendKeys("{Ctrl}v")
    else:
        # 降级：逐字符输入
        _log("[WARN] 剪贴板失败，逐字输入")
        escaped = target.replace("{", "{{").replace("}", "}}")
        search_box.SendKeys(escaped, interval=0.02)

    _rsleep(0.2, 0.3)
    search_box.SendKeys("{Enter}")
    _rsleep(0.7, 1.0)   # 等待搜索结果加载

    return True


def _send_text(wx: "auto.WindowControl", text: str):
    """
    向当前聊天发送文字（剪贴板粘贴）。
    聊天输入框 = 微信主窗口内第 2 个 EditControl（foundIndex=1，0-based）。
    搜索框 = foundIndex=0，聊天输入框 = foundIndex=1（LAVARONG 验证）。
    """
    wx.SetActive()
    _rsleep(0.15, 0.3)

    chat_edit = wx.EditControl(foundIndex=1)
    if not chat_edit.Exists(0, 0):
        raise RuntimeError("未找到聊天输入框（foundIndex=1）")

    chat_edit.Click()
    _rsleep(0.15, 0.3)

    # 剪贴板粘贴（支持中文、换行、特殊字符）
    if _set_clipboard_text(text):
        chat_edit.SendKeys("{Ctrl}v")
        _rsleep(0.15, 0.25)
    else:
        # 降级：转义后 SendKeys（不推荐，特殊字符可能丢失）
        _log("[WARN] 剪贴板失败，使用 SendKeys 降级")
        escaped = text.replace("{", "{{").replace("}", "}}")
        chat_edit.SendKeys(escaped, interval=0.02)
        _rsleep(0.2, 0.3)

    chat_edit.SendKeys("{Enter}")
    _rsleep(0.2, 0.4)
    _log(f"[INFO] 文字已发送: {text[:30]}{'...' if len(text) > 30 else ''}")


def _send_image(wx: "auto.WindowControl", image_path: str):
    """
    向当前聊天发送图片（剪贴板 BMP 粘贴）。
    """
    if not _copy_image_to_clipboard(image_path):
        raise RuntimeError(f"图片写入剪贴板失败: {image_path}")

    wx.SetActive()
    _rsleep(0.15, 0.3)

    chat_edit = wx.EditControl(foundIndex=1)
    if not chat_edit.Exists(0, 0):
        raise RuntimeError("未找到聊天输入框（图片发送）")

    chat_edit.Click()
    _rsleep(0.15, 0.25)
    chat_edit.SendKeys("{Ctrl}v")
    _rsleep(0.4, 0.7)   # 图片粘贴需要更长等待
    chat_edit.SendKeys("{Enter}")
    _rsleep(0.2, 0.4)
    _log(f"[INFO] 图片已发送: {image_path}")


# ─── 统一发送入口（批量任务复用同一个窗口引用）──────────

_cached_wx: Optional["auto.WindowControl"] = None


def call_send(target: str, msg_type: str, text: str, image_path: str):
    """
    统一发送入口（由 cli.py 调用）。
    批量调用时复用 _cached_wx，不重新搜索窗口。
    接口签名与之前版本完全一致。
    """
    global _cached_wx

    if _HAS_WIN32:
        handle = find_window()
        if handle:
            chat = VisualChat(handle)
            chat.open(target, _set_clipboard_text)
            if msg_type in ("文字", "文字+图片"):
                chat.send(target, text, _set_clipboard_text)
            if msg_type in ("图片", "文字+图片"):
                chat.send(target, image_path, _copy_image_to_clipboard)
            return

    # 懒加载 + 有效性检查
    if _cached_wx is None or not _cached_wx.Exists(0, 0):
        _log("[INFO] 初始化/重新获取微信窗口")
        _cached_wx = _get_wx()

    wx = _cached_wx

    # 打开目标会话
    _open_chat(wx, target)

    # 发送内容
    if msg_type == "文字":
        _send_text(wx, text)
    elif msg_type == "图片":
        _send_image(wx, image_path)
    elif msg_type == "文字+图片":
        _send_text(wx, text)
        _rsleep(0.3, 0.5)
        _send_image(wx, image_path)
    else:
        raise ValueError(f"不支持的消息类型: {msg_type}")


# ─── 配置（从外部 config.json 读取）──────────────────────

ROOT = Path(__file__).resolve().parents[1]
CFG_PATH = _real_home() / ".wechat-sender" / "config.json"

SHEET_TASKS = "发送任务"
HEADER_ROW = 2

STATUS_WAITING = "待发送"
STATUS_RUNNING = "发送中"
STATUS_SUCCESS = "发送成功"
STATUS_FAILED = "发送失败"

COL = {
    "seq": "#",
    "app": "* 应用",
    "target": "* 联系人/群聊",
    "msg_type": "* 消息类型",
    "text": "* 文字内容",
    "image": "图片路径",
    "send_time": "发送时间",
    "repeat": "重复",
    "remark": "备注",
    "status": "状态",
}


def load_cfg() -> dict:
    if not CFG_PATH.exists():
        return {}
    with open(CFG_PATH, "r", encoding="utf-8") as f:
        return json.load(f) or {}


@dataclass
class Task:
    row: int
    app: str
    target: str
    msg_type: str
    text: str
    image_path: str
    send_time: Optional[datetime]
    repeat: str
    status: str


def find_columns(ws):
    cols = {}
    for c in range(1, ws.max_column + 1):
        title = (ws.cell(HEADER_ROW, c).value or "").strip()
        if title:
            cols[title] = c
    missing = [v for k, v in COL.items() if v not in cols and not k.startswith("seq")]
    if missing:
        raise RuntimeError(f"模板缺少列: {missing}")
    return cols


def read_tasks(ws, cols) -> list[Task]:
    tasks = []
    for r in range(HEADER_ROW + 1, ws.max_row + 1):
        app = str(ws.cell(r, cols[COL["app"]]).value or "").strip()
        target = str(ws.cell(r, cols[COL["target"]]).value or "").strip()
        msg_type = str(ws.cell(r, cols[COL["msg_type"]]).value or "").strip()
        text = str(ws.cell(r, cols[COL["text"]]).value or "").strip()
        image_path = str(ws.cell(r, cols[COL["image"]]).value or "").strip()
        send_time_raw = ws.cell(r, cols[COL["send_time"]]).value
        repeat_raw = ws.cell(r, cols[COL["repeat"]]).value
        status = str(ws.cell(r, cols[COL["status"]]).value or "").strip()
        if not any([app, target, msg_type, text, image_path, send_time_raw, repeat_raw, status]):
            continue
        send_time = send_time_raw if isinstance(send_time_raw, datetime) else None
        tasks.append(Task(
            row=r, app=app, target=target, msg_type=msg_type,
            text=text, image_path=image_path,
            send_time=send_time,
            repeat=str(repeat_raw).strip() if repeat_raw else "",
            status=status,
        ))
    return tasks


def set_status(ws, cols, row: int, status: str):
    ws.cell(row, cols[COL["status"]]).value = status


def should_send(task: Task, now: datetime) -> bool:
    if task.status.startswith(STATUS_SUCCESS) and not task.repeat:
        return False
    if task.send_time is None:
        return True
    return now >= task.send_time


# ─── 主发送逻辑 ─────────────────────────────────────────

def batch_send(dry_run: bool = False, send_interval: float = 5,
               max_per_minute: int = 8, xlsx_path: str = ""):
    import openpyxl

    if not xlsx_path:
        raise ValueError("未设置 excel_path，请先运行 python app/cli.py setup 配置表格路径")
    xlsx_path = Path(xlsx_path).expanduser()
    if not xlsx_path.exists():
        raise RuntimeError(f"表格不存在: {xlsx_path}")

    print(f"📋 读取表格: {xlsx_path}")

    wb = openpyxl.load_workbook(xlsx_path)
    ws = wb[SHEET_TASKS]
    cols = find_columns(ws)
    tasks = read_tasks(ws, cols)
    now = datetime.now()

    pending = [t for t in tasks if should_send(t, now)]
    if not pending:
        print("⏳ 没有需要发送的任务")
        return

    print(f"📤 开始发送 {len(pending)} 条任务（uiautomation 方案）...")
    if dry_run:
        print("⚠️  模拟运行模式，不会真实发送")

    sent_times: list[datetime] = []
    success_count, fail_count = 0, 0

    for i, task in enumerate(pending):
        window_start = now - timedelta(minutes=1)
        sent_times = [t for t in sent_times if t > window_start]
        if len(sent_times) >= max_per_minute:
            sleep_secs = 60 - (datetime.now() - sent_times[0]).total_seconds()
            if sleep_secs > 0:
                print(f"⏳ 达到每分钟 {max_per_minute} 条限制，等待 {sleep_secs:.0f}s...")
                time.sleep(sleep_secs)
                sent_times = [t for t in sent_times
                              if t > datetime.now() - timedelta(minutes=1)]

        print(f"[{i+1}/{len(pending)}] → {task.target} [{task.msg_type}]", end="")

        set_status(ws, cols, task.row, STATUS_RUNNING)
        wb.save(xlsx_path)

        try:
            if not dry_run:
                if task.app and task.app != "微信":
                    raise ValueError(f"不支持的应用: {task.app}（当前仅支持微信）")
                if not task.target:
                    raise ValueError("联系人/群聊不能为空")
                if task.msg_type not in {"文字", "图片", "文字+图片"}:
                    raise ValueError(f"不支持的消息类型: {task.msg_type}")
                if task.msg_type in {"文字", "文字+图片"} and not task.text:
                    raise ValueError("文字内容不能为空")
                call_send(task.target, task.msg_type, task.text, task.image_path)

            status = f"{STATUS_SUCCESS} {datetime.now().strftime('%H:%M:%S')}"
            set_status(ws, cols, task.row, status)
            sent_times.append(datetime.now())
            success_count += 1
            print(" ✅")
        except Exception as e:
            set_status(ws, cols, task.row, f"{STATUS_FAILED}: {e}")
            fail_count += 1
            print(f" ❌ {e}")

        wb.save(xlsx_path)

        if i < len(pending) - 1 and not dry_run:
            time.sleep(send_interval)

    print(f"\n✅ 完成！成功 {success_count} 条，失败 {fail_count} 条")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="微信批量发送 — Windows 版（uiautomation）")
    parser.add_argument("--dry", action="store_true", help="模拟运行")
    parser.add_argument("--call-single", action="store_true",
                        help="单次发送（由 cli.py call_sender 调用）")
    parser.add_argument("--target", type=str, help="发送目标")
    parser.add_argument("--msg-type", type=str, help="消息类型")
    parser.add_argument("--text", type=str, help="文字内容")
    parser.add_argument("--image-path", type=str, default="", help="图片路径")
    parser.add_argument("--verbose", action="store_true", help="详细日志")
    args = parser.parse_args()

    VERBOSE = args.verbose

    if args.call_single:
        if not args.target or not args.msg_type:
            print("[ERROR] --call-single 需要 --target 和 --msg-type", file=sys.stderr)
            sys.exit(1)
        try:
            call_send(args.target, args.msg_type, args.text or "", args.image_path or "")
            print("✅ 发送成功")
            sys.exit(0)
        except Exception as e:
            print(f"[ERROR] {e}", file=sys.stderr)
            sys.exit(1)

    cfg = load_cfg()
    xlsx_path = cfg.get("excel_path", "")
    send_interval = float(cfg.get("send_interval", 5))
    max_per_minute = int(cfg.get("max_per_minute", 8))
    dry_run = args.dry or cfg.get("dry_run", False)

    batch_send(
        dry_run=dry_run,
        send_interval=send_interval,
        max_per_minute=max_per_minute,
        xlsx_path=xlsx_path,
    )
