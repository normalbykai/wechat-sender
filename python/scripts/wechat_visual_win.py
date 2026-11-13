"""Screen-based sending for the Weixin 4.2 desktop layout."""
import ctypes
import json
import os
from pathlib import Path
import subprocess
from tempfile import TemporaryDirectory
import time

from PIL import ImageChops, ImageGrab, ImageOps
import uiautomation as auto
import win32api
import win32gui
import win32process


def find_window():
    handles = []

    def inspect(handle, _):
        if not win32gui.IsWindowVisible(handle) or win32gui.GetWindowText(handle) != "微信":
            return
        if win32gui.GetClassName(handle) != "Qt51514QWindowIcon":
            return
        pid = win32process.GetWindowThreadProcessId(handle)[1]
        try:
            process = win32api.OpenProcess(0x1000, False, pid)
            try:
                executable = win32process.GetModuleFileNameEx(process, 0)
            finally:
                process.Close()
        except OSError:
            return
        if Path(executable).name.lower() == "weixin.exe":
            handles.append(handle)

    win32gui.EnumWindows(inspect, None)
    if len(handles) > 1:
        raise RuntimeError("检测到多个微信主窗口，请保留一个窗口后重试")
    return handles[0] if handles else None


class VisualChat:
    def __init__(self, handle):
        self.handle = handle
        self.pid = win32process.GetWindowThreadProcessId(handle)[1]
        ctypes.windll.user32.SetProcessDPIAware()
        auto.ControlFromHandle(handle).SetActive()
        if win32gui.GetForegroundWindow() != handle:
            # Windows grants a foreground change after an Alt keystroke.
            win32api.keybd_event(0x12, 0, 0, 0)
            win32api.keybd_event(0x12, 0, 2, 0)
            auto.ControlFromHandle(handle).SetActive()
        time.sleep(0.4)
        self.bounds = win32gui.GetWindowRect(handle)
        self.scale = ctypes.windll.user32.GetDpiForWindow(handle) / 96
        left, top, right, bottom = self.bounds
        self.width, self.height = right - left, bottom - top
        if not self.scale or self.width < self.px(780) or self.height < self.px(500):
            raise RuntimeError("请放大微信窗口后重试")
        screenshot = self.capture()
        y = self.px(72)
        candidates = [
            x for x in range(self.px(250), self.width - self.px(350))
            if all(
                sum(screenshot.getpixel((x + self.px(3), row))[:3])
                - sum(screenshot.getpixel((x - self.px(3), row))[:3]) > 30
                for row in (y, y - self.px(3))
            )
        ]
        if not candidates or candidates[-1] - candidates[0] > self.px(8):
            raise RuntimeError("无法确认微信聊天区位置，请使用默认浅色主题并打开聊天列表")
        self.divider = candidates[len(candidates) // 2]

    def px(self, value):
        return round(value * self.scale)

    def check(self):
        if win32gui.GetForegroundWindow() != self.handle or win32gui.GetWindowRect(self.handle) != self.bounds:
            raise RuntimeError("微信窗口失去焦点或发生移动，本次操作已中止")

    def capture(self):
        self.check()
        return ImageGrab.grab(bbox=self.bounds, all_screens=True)

    def click(self, x, y, popup=False):
        self.check()
        left, top, _, _ = self.bounds
        point = (left + x, top + y)
        hit = win32gui.WindowFromPoint(point)
        if not hit or win32process.GetWindowThreadProcessId(hit)[1] != self.pid:
            raise RuntimeError("微信目标位置被遮挡")
        if not popup and win32gui.GetAncestor(hit, 2) != self.handle:
            raise RuntimeError("微信主窗口上有弹窗")
        auto.Click(*point, waitTime=0.1)

    def keys(self, keys):
        self.check()
        auto.SendKeys(keys, waitTime=0.1)

    def read(self, box, title=False, highlighted=False):
        image = self.capture().crop(box)
        scale = 3 if highlighted else 2
        image = image.resize((image.width * scale, image.height * scale))
        if title:
            image = ImageOps.expand(image.convert("L").point(lambda value: 0 if value < 160 else 255),
                                    border=24, fill=255)
        if highlighted:
            red, green, _ = image.convert("RGB").split()
            image = ImageOps.expand(ImageChops.subtract(green, red).point(lambda value: 0 if value > 85 else 255),
                                    border=24, fill=255)
        with TemporaryDirectory(prefix="wechat-ocr-") as directory:
            path = Path(directory) / "crop.png"
            image.save(path)
            powershell = Path(os.environ["SystemRoot"]) / "System32/WindowsPowerShell/v1.0/powershell.exe"
            result = subprocess.run(
                [str(powershell), "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
                 "-File", str(Path(__file__).with_name("wechat_ocr.ps1")), str(path)],
                capture_output=True, encoding="utf-8", errors="replace",
                creationflags=subprocess.CREATE_NO_WINDOW, timeout=12,
            )
        if result.returncode:
            raise RuntimeError("Windows 屏幕文字识别失败，请检查简体中文 OCR 是否可用")
        return json.loads(result.stdout)

    def verify(self, target):
        px = self.px
        header = self.read((self.divider + px(8), px(39), self.width - px(145), px(69)), title=True)
        if len(header) != 1 or header[0]["text"] != target or header[0]["x"] > 24 + px(36):
            raise RuntimeError("无法通过聊天标题核对收件人，已中止发送")

    def open(self, target, copy_text):
        px = self.px
        self.click(px(130), px(56))
        self.keys("{Ctrl}a")
        if not copy_text(target):
            raise RuntimeError("写入微信搜索内容失败")
        self.keys("{Ctrl}v")
        time.sleep(0.65)
        left, top = px(129), px(76)
        rows = self.read((left, top, self.divider - px(18), self.height - px(12)))
        matches = [
            (row["x"] / 2, (row["y"] + row["h"] / 2) / 2)
            for row in rows if row["text"] == target and px(6) <= row["x"] <= px(40)
        ]
        # WeChat paints matching characters green. Boost their contrast before OCR.
        highlighted = self.read((left, top, min(self.divider - px(18), left + px(180)),
                                 self.height - px(12)), highlighted=True)
        for row in highlighted:
            x = (row["x"] - 24) / 3
            y = (row["y"] - 24 + row["h"] / 2) / 3
            if row["text"] == target and px(3) <= x <= px(20):
                if all(abs(y - match_y) > px(8) for _, match_y in matches):
                    matches.append((x, y))
        if len(matches) != 1:
            self.keys("{Esc}")
            if matches:
                raise RuntimeError("微信搜索结果中有多个匹配项，请使用唯一备注或群名")
            raise RuntimeError("未能识别出与收件人完全一致的微信搜索结果")
        x, y = matches[0]
        self.click(left + round(x) + px(3), top + round(y), popup=True)
        time.sleep(0.45)
        self.verify(target)

    def send(self, target, value, copy):
        px = self.px
        x = self.divider + px(30)
        y = self.height - px(129)
        button = (self.width - px(55), self.height - px(42))
        button_box = (button[0] - px(18), button[1] - px(10),
                      button[0] + px(18), button[1] + px(10))
        self.verify(target)
        if self.send_button_enabled(button_box):
            raise RuntimeError("微信输入区已有草稿，请清空后重试")
        self.click(x, y)
        if not copy(value):
            raise RuntimeError("写入微信发送内容失败")
        self.keys("{Ctrl}v")
        time.sleep(0.45)
        self.verify(target)
        if not self.send_button_enabled(button_box):
            raise RuntimeError("未检测到可用的发送按钮")
        self.click(*button)
        time.sleep(0.45)
        if self.send_button_enabled(button_box):
            raise RuntimeError("微信输入框中仍有内容，请检查是否发送成功")

    def send_button_enabled(self, box):
        image = self.capture().crop(box)
        red, green, _ = image.convert("RGB").split()
        difference = ImageChops.subtract(green, red)
        return sum(difference.histogram()[45:]) > image.width * image.height / 4
