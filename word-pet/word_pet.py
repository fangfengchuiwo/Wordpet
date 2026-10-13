# -*- coding: utf-8 -*-
"""
Word 桌面宠物 (Word Desktop Pet)

功能：
1. 独立启动。后台检测 Word 是否打开，打开后桌宠自动弹出；
   退出 Word 后桌宠自动退出（1 秒缓冲，期间重开 Word 则不退出）。
2. 统计本次 Word 会话中"录入"的数量（默认按字符数：1 个汉字算 1；
   可在 config.json 的 counter.mode 切换为 utf8 / gbk 字节数）。
   中文/英文双通道：汉字上屏即时计入；英文（含输入法拼音组字）
   防抖后计入，互不干扰，总数 = 中文 + 英文。
   删除本次录入的内容会扣减，删除文档原有内容不影响计数。
3. 左键点击桌宠弹出漫画气泡对话框（透明背景），显示固定句式：
   “加油，今天已经工作了？？，写了！！个字。”（？？ = Word 打开时长，！！ = 本次录入数）。
4. 拖动桌宠：靠近（未最大化的）Word 窗口上边缘松手会挂在窗口上边缘，
   并跟随窗口移动，窗口最大化/最小化/关闭时自动脱落；
   靠近屏幕底部会停到任务栏处（稍微沉入一点）；其他位置松手掉落到
   桌面底部，落地即停（无回弹）。动画 GIF 放在 assets/pet/ 下（详见 README）。

依赖：pip install -r requirements.txt
启动：python word_pet.py  （或双击 run.bat）
"""

import ctypes
import ctypes.wintypes as wintypes
import json
import math
import os
import queue
import random
import threading
import time
import tkinter as tk

from PIL import Image, ImageTk, ImageSequence, ImageDraw, ImageFont

# ---------------------------------------------------------------------------
# 可选依赖
# ---------------------------------------------------------------------------
try:
    import pythoncom
    import win32com.client
    HAS_COM = True
except Exception:
    HAS_COM = False

try:
    import psutil
    HAS_PSUTIL = True
except Exception:
    HAS_PSUTIL = False

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE_DIR, "config.json")

DEFAULT_CONFIG = {
    "pet": {
        "size": [140, 140],
        "gifs": {
            "idle": "assets/pet/idle.gif",
            "drag": "assets/pet/drag.gif",
            "hang": "assets/pet/hang.gif",
            "fall": "assets/pet/fall.gif"
        }
    },
    "counter": {
        # 计数单位：chars = 字符数（1 个汉字算 1，默认）
        #           utf8 = UTF-8 字节数（1 个汉字算 3）
        #           gbk  = GBK 字节数（1 个汉字算 2）
        "mode": "chars"
    },
    "dialogs": [
        {
            "title": "打招呼",
            "text": "你好呀！我是你的 Word 桌宠～检测到 Word 已打开，我会一直陪着你写文档！"
        },
        {
            "title": "写累了？",
            "text": "（这里是预留台词。点击桌宠或按“上一句 / 下一句”可切换，直接修改文字会自动保存到 config.json。）"
        },
        {
            "title": "打气",
            "text": "每一个字节都算数！继续加油，写完记得 Ctrl+S 保存～"
        }
    ]
}


def load_config():
    cfg = json.loads(json.dumps(DEFAULT_CONFIG, ensure_ascii=False))  # deep copy
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                user = json.load(f)
            for k, v in user.items():
                if isinstance(v, dict) and isinstance(cfg.get(k), dict):
                    cfg[k].update(v)
                else:
                    cfg[k] = v
        except Exception as e:
            print("读取 config.json 失败，使用默认配置:", e)
    else:
        save_config(cfg)
    return cfg


def save_config(cfg):
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print("保存 config.json 失败:", e)


# ---------------------------------------------------------------------------
# Word 窗口查找（Win32 API，用于把桌宠挂到 Word 窗口边缘）
# ---------------------------------------------------------------------------
_user32 = ctypes.windll.user32


def find_word_windows():
    """枚举所有可见的 Word 主窗口（窗口类 OpusApp），
    返回 [(hwnd, (left, top, right, bottom)), ...]（物理像素）。"""
    results = []

    @ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    def on_window(hwnd, _):
        if _user32.IsWindowVisible(hwnd):
            buf = ctypes.create_unicode_buffer(64)
            _user32.GetClassNameW(hwnd, buf, 64)
            if buf.value == "OpusApp":
                r = wintypes.RECT()
                if _user32.GetWindowRect(hwnd, ctypes.byref(r)):
                    results.append((hwnd, (r.left, r.top, r.right, r.bottom)))
        return True

    _user32.EnumWindows(on_window, 0)
    return results


def window_rect(hwnd):
    """返回窗口 (left, top, right, bottom)，失败返回 None。"""
    r = wintypes.RECT()
    if _user32.GetWindowRect(hwnd, ctypes.byref(r)):
        return (r.left, r.top, r.right, r.bottom)
    return None


def work_area():
    """屏幕工作区 (left, top, right, bottom)——已排除任务栏。失败返回整屏。"""
    r = wintypes.RECT()
    if _user32.SystemParametersInfoW(0x0030, 0, ctypes.byref(r), 0):  # SPI_GETWORKAREA
        return (r.left, r.top, r.right, r.bottom)
    return (0, 0, _user32.GetSystemMetrics(0), _user32.GetSystemMetrics(1))


class _WINDOWPLACEMENT(ctypes.Structure):
    _fields_ = [
        ("length", wintypes.UINT),
        ("flags", wintypes.UINT),
        ("showCmd", wintypes.UINT),
        ("ptMinPosition", wintypes.POINT),
        ("ptMaxPosition", wintypes.POINT),
        ("rcNormalPosition", wintypes.RECT),
    ]


def is_maximized(hwnd):
    """窗口是否处于最大化状态。"""
    placement = _WINDOWPLACEMENT()
    placement.length = ctypes.sizeof(_WINDOWPLACEMENT)
    if _user32.GetWindowPlacement(hwnd, ctypes.byref(placement)):
        return placement.showCmd == 3  # SW_SHOWMAXIMIZED
    return False


# ---------------------------------------------------------------------------
# Word 监控（后台线程）
# ---------------------------------------------------------------------------
class WordMonitor(threading.Thread):
    """轮询 Word：
    - 有 pywin32：通过 COM 拿到 ActiveDocument，统计“本次录入”的数量
      （单位由 counter.mode 决定：字符数 / UTF-8 字节 / GBK 字节，默认字符数）。
    - 中文/英文双通道计数：
      * 中文通道（汉字/全角字符）：即时结算。汉字上屏直接计入。
        输入法组字的拼音属于英文通道，不会影响中文计数。
      * 英文通道（字母/半角字符）：防抖结算。英文部分连续 EN_SETTLE_SEC
        秒不变才计入，瞬时的拼音组字中间态会被自动跳过。
        中文输入不会重置英文通道的防抖计时，两通道完全独立。
      * 总计数 = 中文 + 英文。用英文输入法打字时全部走英文通道（即原防抖方案）。
    - 删除“本次录入”的内容 → 计数扣减（两通道各自独立归属）；
      删除“文档原有内容”（基线以下）→ 计数不受影响。
    - 无 pywin32：用 psutil 检测 WINWORD.EXE 进程（只做打开检测，不统计）。
    """

    POLL_INTERVAL = 0.5
    EN_SETTLE_SEC = 1.2          # 英文通道防抖：英文内容稳定这么久才结算
    UNITS = {"chars": "字", "utf8": "字节(UTF-8)", "gbk": "字节(GBK)"}

    def __init__(self, events, count_mode="chars"):
        super().__init__(daemon=True)
        self.events = events          # queue.Queue, "open" / "close"
        self.count_mode = count_mode if count_mode in self.UNITS else "chars"
        self.word_open = False
        self.cn_session = 0           # 中文通道：本次会话累计录入
        self.en_session = 0           # 英文通道：本次会话累计录入
        self.note = ""                # 给对话框显示的提示
        self._doc_key = None          # 当前统计的文档标识
        self._base_cn = 0             # 基线：文档检测到时的中文数量（原有内容）
        self._base_en = 0             # 基线：文档检测到时的英文数量
        self._settled_cn = 0          # 中文通道已结算数量（即时更新）
        self._settled_en = 0          # 英文通道已结算数量（防抖后更新）
        self._last_seen_en = 0        # 英文通道最近读到的数量（可能是拼音中间态）
        self._last_en_change = 0.0    # 英文通道上次变化的时间戳
        self._last_count = None       # 上次 Characters.Count（用于变化检测，避免频繁读全文）
        self.open_since = None        # Word 本次打开的时间戳（关闭后清空）

    @property
    def session_bytes(self):
        """总计数 = 中文 + 英文。"""
        return self.cn_session + self.en_session

    @property
    def unit(self):
        return self.UNITS[self.count_mode]

    @property
    def elapsed_str(self):
        """Word 本次已打开的时长，格式如“1小时5分”“3分20秒”。"""
        if self.open_since is None:
            return "0秒"
        secs = max(0, int(time.time() - self.open_since))
        h, m, s = secs // 3600, secs % 3600 // 60, secs % 60
        if h:
            return f"{h}小时{m}分"
        if m:
            return f"{m}分{s}秒"
        return f"{s}秒"

    # ---- 线程主体 -----------------------------------------------------
    def run(self):
        if HAS_COM:
            try:
                pythoncom.CoInitialize()
            except Exception:
                pass
        while True:
            try:
                self._poll()
            except Exception:
                pass
            time.sleep(self.POLL_INTERVAL)

    # ---- 轮询 ---------------------------------------------------------
    def _poll(self):
        if HAS_COM:
            word = None
            try:
                word = win32com.client.GetActiveObject("Word.Application")
            except Exception:
                word = None
            if word is not None:
                self.note = ""
                self._set_open(True)
                try:
                    doc = word.ActiveDocument
                except Exception:
                    doc = None
                if doc is not None:
                    self._count(doc)
                return

        if HAS_PSUTIL:
            found = False
            for p in psutil.process_iter(["name"]):
                name = (p.info.get("name") or "").lower()
                if name == "winword.exe":
                    found = True
                    break
            if not HAS_COM:
                self.note = "未安装 pywin32，无法统计输入（pip install pywin32）"
            elif found:
                # Word 进程在，但 COM 连不上（Word 刚启动还没在系统注册，
                # 切换一次窗口焦点后即可正常统计）
                self.note = "已检测到 Word，等待连接后开始统计（点击一下其他窗口即可）"
            else:
                self.note = ""
            self._set_open(found)
        else:
            self.note = "未安装 pywin32 / psutil，仅显示桌宠，无法统计输入"
            self._set_open(True)

    # ---- 状态切换 ------------------------------------------------------
    def _set_open(self, open_):
        if open_ != self.word_open:
            self.word_open = open_
            if open_:
                self.open_since = time.time()
                self.events.put("open")
            else:
                self.open_since = None
                self.events.put("close")
                # Word 关闭：会话结束，清零
                self.cn_session = 0
                self.en_session = 0
                self._doc_key = None
                self._base_cn = 0
                self._base_en = 0
                self._settled_cn = 0
                self._settled_en = 0
                self._last_seen_en = 0
                self._last_en_change = 0.0
                self._last_count = None

    # ---- 字节统计 ------------------------------------------------------
    def _count(self, doc):
        """双通道统计规则：
        - 中文通道（汉字/全角）：即时结算，汉字上屏立刻计入；
        - 英文通道（字母/半角）：防抖结算，英文内容稳定 EN_SETTLE_SEC 秒才计入，
          输入法组字的拼音中间态会被自动跳过，中文输入不影响英文计时；
        - 新增 → 累加；删除“本次录入” → 扣减（各通道独立归属）；
          删除“文档原有内容”（基线以下）→ 不扣减。
        """
        try:
            key = doc.FullName
        except Exception:
            key = str(id(doc))
        try:
            count = int(doc.Characters.Count)
        except Exception:
            return

        if key != self._doc_key:
            # 新文档（首次检测到 / 切换文档）：以当前内容为基线重新计数
            self._doc_key = key
            cn, en = self._doc_cn_en(doc)
            self._base_cn = cn
            self._base_en = en
            self._settled_cn = cn
            self._settled_en = en
            self._last_seen_en = en
            self._last_en_change = time.time()
            self._last_count = count
            self.cn_session = 0
            self.en_session = 0
            return

        # 英文通道有待定值（拼音/连续输入未结算）时，即使 Characters.Count
        # 没变也要再读一次，用于确认英文部分已稳定。
        pending_en = self._last_seen_en != self._settled_en
        if count == self._last_count and not pending_en:
            # 内容没变，无需读取
            return

        cn, en = self._doc_cn_en(doc)
        self._last_count = count
        now = time.time()

        # --- 中文通道：即时结算 ---
        # 拼音组字只影响英文通道，cn 不变；汉字上屏即终态，直接计入。
        if cn != self._settled_cn:
            self._apply_delta("cn", self._settled_cn, cn, self._base_cn)
            self._settled_cn = cn

        # --- 英文通道：防抖结算 ---
        if en != self._last_seen_en:
            # 英文内容仍在变化（拼音组字中 / 连续输入中），暂不结算
            self._last_seen_en = en
            self._last_en_change = now
        elif (en != self._settled_en
              and now - self._last_en_change >= self.EN_SETTLE_SEC):
            # 英文内容已稳定足够久，结算
            self._apply_delta("en", self._settled_en, en, self._base_en)
            self._settled_en = en

    def _apply_delta(self, channel, prev, cur, base):
        """把一次结算的变化（prev → cur）计入对应通道（'cn' / 'en'）。"""
        attr = channel + "_session"
        session = getattr(self, attr)
        delta = cur - prev
        if delta > 0:
            # 新录入（输入 / 粘贴）：累加
            setattr(self, attr, session + delta)
        elif delta < 0:
            # 删除：只扣减可归属为“本次录入”的部分。
            # 高于基线的数量才可能是本次录入的；删除量中超出
            # “基线以上数量”的部分一定来自原有内容，不扣减。
            above_base = max(0, prev - base)
            typed_removed = min(-delta, above_base, session)
            setattr(self, attr, session - typed_removed)

    def _doc_cn_en(self, doc):
        """统计文档内容，返回 (中文数量, 英文数量)。
        中文 = 汉字及全角字符（Unicode >= U+2E80）；英文 = 其余（字母/数字/半角符号）。
        数量按 counter.mode 计权：chars 1/1；gbk 2/1；utf8 3/1。
        """
        try:
            text = doc.Content.Text
            if self.count_mode == "utf8":
                cn_w, en_w = 3, 1
            elif self.count_mode == "gbk":
                cn_w, en_w = 2, 1
            else:
                cn_w, en_w = 1, 1
            en_chars = 0
            for ch in text:
                if ch < "　":  # U+3000 以下为半角/拉丁区
                    en_chars += 1
            cn_chars = len(text) - en_chars
            return cn_chars * cn_w, en_chars * en_w
        except Exception:
            # 兜底：全部按英文（半角）估算
            try:
                n = int(doc.Characters.Count)
                return 0, n
            except Exception:
                return 0, 0


# ---------------------------------------------------------------------------
# 占位动画（未放 GIF 时自动生成一只小圆宠）
# ---------------------------------------------------------------------------
def _load_font(size=13):
    for p in (
        r"C:\Windows\Fonts\msyh.ttc",
        r"C:\Windows\Fonts\simhei.ttf",
        r"C:\Windows\Fonts\arial.ttf",
    ):
        try:
            return ImageFont.truetype(p, size)
        except Exception:
            continue
    return ImageFont.load_default()


def _placeholder_frames(state, size):
    """没有 GIF 文件时，用 PIL 画一只会眨眼/上下浮动的小圆宠。"""
    w, h = size
    palette = {
        "idle": (255, 205, 100, 255),
        "drag": (239, 100, 130, 255),
        "hang": (80, 210, 170, 255),
        "fall": (110, 170, 235, 255),
    }
    body = palette.get(state, palette["idle"])
    outline = (70, 60, 60, 255)
    label = {"idle": "待机", "drag": "拖动中", "hang": "挂在边缘", "fall": "下落中"}.get(state, "")
    font = _load_font(13)

    frames = []
    n = 10
    for i in range(n):
        phase = 2 * math.pi * i / n
        img = Image.new("RGBA", (w, h), (0, 0, 0, 0))
        d = ImageDraw.Draw(img)
        bob = math.sin(phase) * (6 if state == "idle" else 3)
        cx, cy = w // 2, int(h * 0.42) + int(bob)
        r = int(min(w, h) * 0.30)

        # 耳朵
        d.ellipse([cx - r * 0.75, cy - r * 1.15, cx - r * 0.15, cy - r * 0.45],
                  fill=body, outline=outline, width=2)
        d.ellipse([cx + r * 0.15, cy - r * 1.15, cx + r * 0.75, cy - r * 0.45],
                  fill=body, outline=outline, width=2)
        # 身体
        d.ellipse([cx - r, cy - r, cx + r, cy + r], fill=body, outline=outline, width=2)
        # 眼睛
        blink = (i % 5 == 4) or state == "hang"
        if blink:
            d.line([cx - r * 0.50, cy - r * 0.10, cx - r * 0.20, cy - r * 0.10], fill=outline, width=3)
            d.line([cx + r * 0.20, cy - r * 0.10, cx + r * 0.50, cy - r * 0.10], fill=outline, width=3)
        else:
            d.ellipse([cx - r * 0.50, cy - r * 0.35, cx - r * 0.20, cy - r * 0.05],
                      fill=(255, 255, 255, 255), outline=outline, width=2)
            d.ellipse([cx + r * 0.20, cy - r * 0.35, cx + r * 0.50, cy - r * 0.05],
                      fill=(255, 255, 255, 255), outline=outline, width=2)
            d.ellipse([cx - r * 0.42, cy - r * 0.30, cx - r * 0.26, cy - r * 0.12], fill=outline)
            d.ellipse([cx + r * 0.26, cy - r * 0.30, cx + r * 0.42, cy - r * 0.12], fill=outline)
        # 嘴
        d.arc([cx - r * 0.30, cy - r * 0.10, cx + r * 0.30, cy + r * 0.35],
              start=20, end=160, fill=outline, width=2)
        # 底部提示
        text = f"{label}·可换GIF"
        bbox = d.textbbox((0, 0), text, font=font)
        tw = bbox[2] - bbox[0]
        d.text(((w - tw) // 2, h - 22), text, font=font, fill=(100, 100, 100, 255))
        frames.append(img)
    return frames


# ---------------------------------------------------------------------------
# GIF 解码缓存（大 GIF 每次重新解码非常卡，这里只解码一次）
# ---------------------------------------------------------------------------
_GIF_DECODED = {}    # (绝对路径, (w,h), zoom, shift_y) -> [(PIL.Image, dur), ...]
_GIF_PHOTOS = {}     # 同上 key -> [(PhotoImage, dur), ...]
_GIF_RGBA = {}       # 同上 key -> [(PIL.Image(RGBA), dur), ...]（抖动缩放动画用）
_GIF_LOCK = threading.Lock()


def decode_gif_frames(path, size, zoom=1.0, shift_y=0):
    """解码 GIF 并缩放到 size（纯 PIL 操作，可在后台线程运行），结果全局缓存。
    zoom：整体放大倍数（大于 1 时先放大再从中心裁回 size，人物显得更大）；
    shift_y：内容整体向下平移的像素数（负数为向上）。
    """
    key = (path, tuple(size), zoom, shift_y)
    with _GIF_LOCK:
        if key in _GIF_DECODED:
            return _GIF_DECODED[key]
    w, h = int(size[0]), int(size[1])
    frames = []
    rgba_frames = []   # 带 alpha 的版本（抖动缩放动画再缩放时不产生黑框）
    try:
        im = Image.open(path)
        for fr in ImageSequence.Iterator(im):
            fr = fr.convert("RGBA")
            if zoom != 1.0:
                # 整帧放大（不裁剪）——窗口会跟着变大，人物完整显示
                zw, zh = max(1, int(w * zoom)), max(1, int(h * zoom))
                fr = fr.resize((zw, zh), Image.LANCZOS)
            elif fr.size != (w, h):
                fr = fr.resize((w, h), Image.LANCZOS)
            if shift_y:
                # 整体向下平移
                canvas = Image.new("RGBA", fr.size, (0, 0, 0, 0))
                canvas.paste(fr, (0, shift_y))
                fr = canvas
            if fr.mode == "RGBA":
                if zoom != 1.0 or fr.size != (w, h):
                    fr = _defringe(fr)   # 缩放混入的黑色暗边 → 换成邻近颜色
                rgba_frames.append((fr, max(30, int(fr.info.get("duration", 100) or 100))))
                fr = _flatten_transparent(fr)   # 压平 alpha，避免 Tk 与底色混合发暗
            dur = int(fr.info.get("duration", 100) or 100)
            frames.append((fr, max(30, dur)))
    except Exception as e:
        print("加载 GIF 失败:", path, e)
    with _GIF_LOCK:
        _GIF_DECODED[key] = frames
        if rgba_frames:
            _GIF_RGBA[key] = rgba_frames
    return frames


def _defringe(img, radius=3):
    """去掉缩放产生的黑色细边：LANCZOS 插值会把透明像素的黑色 RGB (0,0,0)
    混进边缘，形成半透明暗边。把这类像素的 RGB 换成邻近不透明像素的颜色。"""
    w, h = img.size
    px = img.load()
    semi = [(x, y) for y in range(h) for x in range(w)
            if 0 < px[x, y][3] < 250]
    for x, y in semi:
        best = None
        for r in range(1, radius + 1):
            found = [px[x + dx, y + dy][:3]
                     for dy in range(-r, r + 1) for dx in range(-r, r + 1)
                     if 0 <= x + dx < w and 0 <= y + dy < h
                     and px[x + dx, y + dy][3] >= 250]
            if found:
                n = len(found)
                best = tuple(sum(c[i] for c in found) // n for i in range(3))
                break
        if best:
            px[x, y] = (best[0], best[1], best[2], px[x, y][3])
    return img


PET_TRANS = (1, 2, 3)   # 窗口透明色（与 PetWindow.TRANSPARENT 一致）


def _flatten_transparent(img, trans=PET_TRANS, thresh=128):
    """把 RGBA 帧压平成 RGB：透明像素填成窗口透明色。
    Tk 显示半透明像素时会与控件背景色（近黑透明色）混合，人物边缘
    会被染黑一圈，所以直接按阈值二值化 alpha，配合 _defringe 保证边缘干净。"""
    alpha = img.getchannel("A")
    mask = alpha.point(lambda v: 255 if v >= thresh else 0)
    out = Image.new("RGB", img.size, trans)
    out.paste(img.convert("RGB"), (0, 0), mask)
    return out


def gif_photos(path, size, zoom=1.0, shift_y=0):
    """取某 GIF 的 PhotoImage 帧列表（带缓存；PhotoImage 只能在主线程建）。
    第一次调用会把解码缓存消费掉以省内存。"""
    key = (path, tuple(size), zoom, shift_y)
    photos = _GIF_PHOTOS.get(key)
    if photos is not None:
        return photos
    with _GIF_LOCK:
        decoded = _GIF_DECODED.pop(key, None)
    if not decoded:
        return []
    photos = [(ImageTk.PhotoImage(im), dur) for im, dur in decoded]
    _GIF_PHOTOS[key] = photos
    return photos


def _gif_params(cfg, state):
    """某状态的渲染参数：整体下移量 + 放大倍数（均可按状态配置）。"""
    pet_cfg = cfg.get("pet", {})
    sy = pet_cfg.get("shift_y", 0)
    if isinstance(sy, dict):          # 按状态单独设置：{"idle": 24, ...}
        shift_y = int(sy.get(state, 0))
    else:
        shift_y = int(sy)
    zoom = float(pet_cfg.get("zoom", {}).get(state, 1.0))
    return zoom, shift_y


def preload_gifs(cfg):
    """后台线程预解码所有状态的 GIF，让状态切换零等待。"""
    size = tuple(cfg["pet"].get("size", [140, 140]))

    def worker():
        for state, rel in cfg["pet"].get("gifs", {}).items():
            p = rel if os.path.isabs(rel) else os.path.join(BASE_DIR, rel)
            if os.path.exists(p):
                zoom, shift_y = _gif_params(cfg, state)
                try:
                    decode_gif_frames(p, size, zoom, shift_y)
                except Exception:
                    pass

    threading.Thread(target=worker, daemon=True).start()


# ---------------------------------------------------------------------------
# GIF 播放器
# ---------------------------------------------------------------------------
class GifPlayer:
    def __init__(self, label):
        self.label = label
        self.frames = []       # [(PhotoImage, duration_ms), ...]
        self.pil_frames = []   # 对应的 PIL 帧（抖动缩放动画用）
        self.idx = 0
        self._token = 0
        self._job = None

    def play(self, rel_path, fallback_state, size, zoom=1.0, shift_y=0):
        self._token += 1
        token = self._token
        if self._job:
            try:
                self.label.after_cancel(self._job)
            except Exception:
                pass
            self._job = None

        frames = []
        pil_frames = []
        if rel_path:
            p = rel_path if os.path.isabs(rel_path) else os.path.join(BASE_DIR, rel_path)
            if os.path.exists(p):
                frames = gif_photos(p, size, zoom, shift_y)   # 命中缓存 → 瞬间切换
                if not frames:
                    # 首次播放且后台还没解码完：同步解码一次
                    decode_gif_frames(p, size, zoom, shift_y)
                    frames = gif_photos(p, size, zoom, shift_y)
                pil_frames = _GIF_RGBA.get((p, tuple(size), zoom, shift_y), [])

        if not frames:
            pl = _placeholder_frames(fallback_state, size)
            frames = [(ImageTk.PhotoImage(f), 90) for f in pl]
            pil_frames = [(f, 90) for f in pl]

        self.frames = frames
        self.pil_frames = pil_frames
        self.idx = 0
        self._step(token)

    def _step(self, token):
        if token != self._token:
            return
        img, dur = self.frames[self.idx]
        self.label.config(image=img)
        self.idx = (self.idx + 1) % len(self.frames)
        self._job = self.label.after(dur, self._step, token)

    def pause(self):
        if self._job:
            try:
                self.label.after_cancel(self._job)
            except Exception:
                pass
            self._job = None

    def resume(self):
        if not self._job and self.frames:
            self._step(self._token)


# ---------------------------------------------------------------------------
# 对话框（点击桌宠弹出，可编辑台词 + 实时字节统计）
# ---------------------------------------------------------------------------
class PetDialog:
    """方形对话气泡（透明背景、圆角矩形 + 尾巴指向桌宠，自动避让屏幕边缘）。"""

    W, H = 300, 150          # 气泡窗口大小
    TRANS = "#010203"        # 透明色（与桌宠窗口一致）
    FILL = (255, 253, 245)   # 气泡底色
    OUTLINE = (25, 25, 25)   # 气泡描边

    # 两种台词句式，每次弹出随机选一种（{count}+{unit}=字数，{elapsed}=用时）
    TEMPLATES = [
        "加油，今天已经工作了{elapsed}，写了{count}{unit}。",
        "厉害，竟然已经写了{count}{unit}，仅用时{elapsed}。",
    ]

    def __init__(self, pet, auto_close_ms=3000):
        self.pet = pet
        self.cfg = pet.cfg
        self.template = random.choice(self.TEMPLATES)

        self.top = tk.Toplevel(pet.win)
        self.top.overrideredirect(True)            # 无边框
        self.top.attributes("-topmost", True)
        self.top.configure(bg=self.TRANS)
        try:
            self.top.attributes("-transparentcolor", self.TRANS)
        except Exception:
            pass

        self.canvas = tk.Canvas(self.top, width=self.W, height=self.H,
                                bg=self.TRANS, highlightthickness=0, bd=0,
                                takefocus=0)
        self.canvas.pack()
        self._bubble_img = self._draw_bubble()
        self.canvas.create_image(0, 0, image=self._bubble_img, anchor="nw")

        self._place_widgets()
        self.reposition()                          # 根据桌宠位置确定出现位置

        # 气泡可拖动（空白处按住拖动）
        self.canvas.bind("<Button-1>", self._on_canvas_press)
        self.canvas.bind("<B1-Motion>", self._on_canvas_drag)

        self._tick()
        self._auto_close_job = self.top.after(auto_close_ms, self.close)

        # 位置重叠时，桌宠显示在气泡上层
        self.pet.win.lift(self.top)

    # ---- 气泡绘制 --------------------------------------------------------
    def _draw_bubble(self):
        """用 PIL 画圆角矩形气泡（无尾巴），返回 PhotoImage。"""
        S = 2                                       # 2 倍超采样抗锯齿
        img = Image.new("RGB", (self.W * S, self.H * S), (1, 2, 3))
        d = ImageDraw.Draw(img)
        top, bottom = 6 * S, (self.H - 6) * S
        # 外描边 + 内填充（内缩 5px 形成描边厚度）
        d.rounded_rectangle([6 * S, top, (self.W - 6) * S, bottom],
                            radius=18 * S, fill=self.OUTLINE)
        d.rounded_rectangle([11 * S, top + 5 * S, (self.W - 11) * S,
                             bottom - 5 * S], radius=14 * S, fill=self.FILL)
        img = img.resize((self.W, self.H), Image.LANCZOS)
        return ImageTk.PhotoImage(img)

    def _place_widgets(self):
        c = self.canvas
        # 固定句式：加油，今天已经工作了？？，写了！！个字。
        # ？？ = Word 本次打开时长，！！ = 本次录入的字数
        self._sentence_id = c.create_text(
            self.W // 2, 54, text="加油，今天已经工作了0秒，写了0个字。",
            font=("Microsoft YaHei UI", 13, "bold"), fill="#333333",
            width=260, justify="center")

        # 提示信息（如“等待连接”）
        self._note_id = c.create_text(
            self.W // 2, 104, text="", font=("Microsoft YaHei UI", 9),
            fill="#c0392b", width=260, justify="center", anchor="center")

    # ---- 定位（根据桌宠位置，自动避让屏幕边缘和 Word 窗口）----------------
    def reposition(self):
        if not self.top.winfo_exists():
            return
        pet = self.pet
        px, py = pet.pos()
        pw, ph = pet.size
        sw = self.top.winfo_screenwidth()
        sh = self.top.winfo_screenheight()
        margin = 6

        def clamp(x, y):
            x = int(min(max(4, x), sw - self.W - 4))
            y = int(min(max(4, y), sh - self.H - 4))
            return x, y

        # 候选位置：桌宠上方（优先）、下方、右侧、左侧
        cands = [
            (px + pw / 2 - self.W / 2, py - self.H - margin),
            (px + pw / 2 - self.W / 2, py + ph + margin),
            (px + pw + margin, py + ph / 2 - self.H / 2),
            (px - self.W - margin, py + ph / 2 - self.H / 2),
        ]

        # 需要避开的区域：Word 窗口（未最小化的）——尽量不挡住 Word 内容。
        # 桌宠自身不用避让：重叠时桌宠始终显示在气泡上层。
        word_rects = []
        try:
            for _hwnd, (l, t, r, b) in find_word_windows():
                if l > -10000:
                    word_rects.append((l, t, r, b))
        except Exception:
            pass

        def hits_word(x, y):
            l, t, r, b = x, y, x + self.W, y + self.H
            return any(l < wr and r > wl and t < wb and b > wt
                       for wl, wt, wr, wb in word_rects)

        # 桌宠在屏幕底部 / 拖动中：不避让 Word 窗口，气泡直接放在桌宠正上方，
        # 保证完整显示（只会盖住 Word 窗口底部一点）
        at_bottom = (py + ph) >= sh - 100
        dragging = getattr(self.pet, "_press", None) is not None
        if at_bottom or dragging:
            pos = clamp(cands[0][0], cands[0][1])
        else:
            pos = None
            for cx_, cy_ in cands:
                tx, ty = clamp(cx_, cy_)
                if not hits_word(tx, ty):
                    pos = (tx, ty)
                    break
            if pos is None:
                # 四个方向都会挡住（如 Word 最大化）→ 退回桌宠上方
                pos = clamp(cands[0][0], cands[0][1])
        self.top.geometry(f"+{pos[0]}+{pos[1]}")

    # ---- 气泡拖动 --------------------------------------------------------
    def _on_canvas_press(self, e):
        self._drag_pos = (e.x_root - self.top.winfo_x(),
                          e.y_root - self.top.winfo_y())

    def _on_canvas_drag(self, e):
        if getattr(self, "_drag_pos", None):
            self.top.geometry(f"+{e.x_root - self._drag_pos[0]}"
                              f"+{e.y_root - self._drag_pos[1]}")

    # ---- 实时刷新 --------------------------------------------------------
    def _tick(self):
        if not self.top.winfo_exists():
            return
        m = self.pet.monitor
        self.canvas.itemconfig(
            self._sentence_id,
            text=self.template.format(elapsed=m.elapsed_str,
                                      count=m.session_bytes, unit=m.unit))
        note = m.note or ""
        self.canvas.itemconfig(self._note_id, text=note)
        # 垂直居中：无提示信息时句子在气泡正中；有提示时句子上移、提示在下
        if note:
            self.canvas.coords(self._sentence_id, self.W // 2, 52)
            self.canvas.coords(self._note_id, self.W // 2, 104)
        else:
            self.canvas.coords(self._sentence_id, self.W // 2, self.H // 2)
        self.top.after(500, self._tick)

    def close(self):
        if getattr(self, "_auto_close_job", None):
            try:
                self.top.after_cancel(self._auto_close_job)
            except Exception:
                pass
            self._auto_close_job = None
        self.pet.on_dialog_closed(self)
        try:
            self.top.destroy()
        except Exception:
            pass


# ---------------------------------------------------------------------------
# 桌宠窗口
# ---------------------------------------------------------------------------
class PetWindow:
    SNAP_DIST = 100         # 距边缘多少像素以内算“挂住”
    SNAP_ABOVE = 250        # Word 窗口边缘外侧（窗口外）的吸附范围，更宽松
    TASKBAR_SINK = 20       # 底部落位时沉入任务栏的像素（贴任务栏再稍微下面一点）
    WORD_EDGE_OVERLAP = 12  # 挂在 Word 窗口边缘时，嵌入边缘的像素（视觉上“扒住”）
    FOLLOW_INTERVAL = 16    # 挂靠 Word 窗口时的跟随刷新间隔（毫秒，约 60fps）
    BOTTOM_SNAP_DIST = 130  # 屏幕底部的吸附范围（略大于窗口吸附）

    def __init__(self, root, cfg, monitor):
        self.root = root
        self.cfg = cfg
        self.monitor = monitor
        self.base_size = tuple(cfg["pet"].get("size", [140, 140]))
        self._cur_size = self.base_size   # 当前窗口大小（拖动放大时会变大）
        self._pos = None                  # 自记录的窗口位置（外部移动用）
        self._hwnd = None                 # 顶层窗口系统句柄（懒加载）

        self.state = "hidden"
        self.visible = False
        self.dialog = None
        self.dialog_index = 0

        self._press = None
        self._moved = False
        self._drag_offset = (0, 0)
        self._pending_pos = None   # 拖动待生效位置（合并高频移动事件用）
        self._move_job = None
        self._last_move_t = 0.0
        self._vy = 0.0
        self._anim_token = 0
        self._last_click = time.time()   # 上次点击桌宠的时间（用于闲置提醒）
        self._attach = None         # 挂靠信息 (hwnd, edge)，挂 Word 窗口边缘时非空
        self._attach_offset = 0     # 沿边缘方向的偏移（跟随窗口移动时保持）
        self._follow_job = None
        self._press_attach = None   # 按下前的挂靠状态（单击后恢复）

        self.TRANSPARENT = "#010203"
        self.win = tk.Toplevel(root)
        self.win.withdraw()
        self.win.overrideredirect(True)
        self.win.attributes("-topmost", True)
        self.win.config(bg=self.TRANSPARENT)
        try:
            self.win.attributes("-transparentcolor", self.TRANSPARENT)
        except Exception:
            pass

        self.label = tk.Label(self.win, bg=self.TRANSPARENT, cursor="fleur",
                              highlightthickness=0, takefocus=0)
        self.label.pack()
        self.player = GifPlayer(self.label)

        self.label.bind("<Button-1>", self._on_press)
        self.label.bind("<B1-Motion>", self._on_drag)
        self.label.bind("<ButtonRelease-1>", self._on_release)
        self.label.bind("<Button-3>", self._on_menu)

        # 闲置提醒：每 1 分钟检查一次，5 分钟没点桌宠就自动弹气泡 1 秒
        self.win.after(60000, self._check_idle)

    # ---- 屏幕坐标 ------------------------------------------------------
    @property
    def size(self):
        return self._cur_size

    def _resize_to(self, wh, keep_center=True):
        """调整窗口大小（状态缩放用），默认保持窗口中心不动。"""
        w1, h1 = int(wh[0]), int(wh[1])
        if (w1, h1) == self._cur_size:
            return
        if keep_center and self.visible:
            cx = self.pos()[0] + self._cur_size[0] // 2
            cy = self.pos()[1] + self._cur_size[1] // 2
            x, y = cx - w1 // 2, cy - h1 // 2
        else:
            x, y = self.pos()
        x = max(0, min(x, self.screen_w - w1))
        y = max(0, min(y, self.screen_h - h1))
        self._cur_size = (w1, h1)
        self._pos = (int(x), int(y))
        self.win.geometry(f"{w1}x{h1}+{int(x)}+{int(y)}")

    def _default_size(self):
        """常规（待机）状态下的窗口大小。"""
        zoom, _ = _gif_params(self.cfg, "idle")
        return (int(self.base_size[0] * zoom), int(self.base_size[1] * zoom))

    @property
    def screen_w(self):
        return self.win.winfo_screenwidth()

    @property
    def screen_h(self):
        return self.win.winfo_screenheight()

    def bottom_y(self):
        """底部固定落位：贴到任务栏再稍微沉进去一点。"""
        return int(work_area()[3]) + self.TASKBAR_SINK - self.size[1]

    def clamp_pos(self, x, y):
        x = max(0, min(x, self.screen_w - self.size[0]))
        y = max(0, min(y, self.screen_h - self.size[1]))
        return int(x), int(y)

    def move_to(self, x, y):
        x, y = self.clamp_pos(x, y)
        self._pos = (int(x), int(y))
        hwnd = self._top_hwnd()
        if hwnd:
            # 用系统接口直接移动窗口，绕开 Tk 的 geometry 处理管线：
            # 快速拖动时 Tk 处理 geometry 会同步重绘，透明底色
            # 来不及混合就闪出黑框；SetWindowPos 只移动合成层，不触发重绘
            _user32.SetWindowPos(hwnd, 0, int(x), int(y), 0, 0,
                                 0x0001 | 0x0004 | 0x0010)  # NOSIZE|NOZORDER|NOACTIVATE
        else:
            self.win.geometry(f"+{int(x)}+{int(y)}")
        # 桌宠移动时，打开中的对话气泡随动
        if self.dialog is not None and self.dialog.top.winfo_exists():
            self.dialog.reposition()

    def pos(self):
        """当前窗口位置。外部移动后 winfo_x 可能没来得及刷新，以自记录为准。"""
        if self._pos is not None:
            return self._pos
        return (self.win.winfo_x(), self.win.winfo_y())

    def _top_hwnd(self):
        """桌宠顶层窗口的系统句柄（Tk 会把 Toplevel 再包一层 wrapper 窗口）。"""
        if self._hwnd is None:
            try:
                h = self.win.winfo_id()
                p = _user32.GetParent(h)
                self._hwnd = p if p else h
            except Exception:
                self._hwnd = 0
        return self._hwnd

    def cancel_anim(self):
        self._anim_token += 1

    # ---- 状态 / 动画 ----------------------------------------------------
    def set_state(self, s):
        if s == self.state:
            return
        self.state = s
        if s == "hidden":
            return
        # 按该状态的 zoom 缩放窗口（拖动放大 → 整只完整显示）
        zoom, shift_y = _gif_params(self.cfg, s)
        self._resize_to((int(self.base_size[0] * zoom),
                         int(self.base_size[1] * zoom)))
        path = self.cfg["pet"].get("gifs", {}).get(s)
        self.player.play(path, s, self.base_size, zoom, shift_y)

    def show(self):
        """Word 打开：弹出（从屏幕底部升起）。"""
        if self.visible:
            return
        self.visible = True
        self.cancel_anim()
        token = self._anim_token
        # 先切到待机状态（窗口 140→待机放大后的尺寸），再算落位：
        # 否则 bottom_y() 按 140px 算，升起终点会比落底位置低一截，
        # 导致“起始位置”和“落下位置”不一致
        self.set_state("idle")
        x = (self.screen_w - self.size[0]) // 2
        y0 = self.screen_h + 10
        target = self.bottom_y()
        self.win.geometry(f"{self.size[0]}x{self.size[1]}+{x}+{y0}")
        self.win.deiconify()
        self._rise(token, x, y0, target, 0)

    def _rise(self, token, x, y0, target, i):
        if token != self._anim_token or not self.visible:
            return
        steps = 16
        if i > steps:
            return
        t = i / steps
        ease = 1 - (1 - t) ** 3
        y = int(y0 + (target - y0) * ease)
        self.move_to(x, y)
        self.win.after(16, self._rise, token, x, y0, target, i + 1)

    def hide(self):
        if not self.visible:
            return
        self.visible = False
        self.cancel_anim()
        self._detach()
        if self.dialog is not None:
            try:
                self.dialog.close()
            except Exception:
                pass
        self.win.withdraw()
        self.state = "hidden"

    # ---- 拖拽 ----------------------------------------------------------
    def _on_press(self, e):
        self._last_click = time.time()
        self.cancel_anim()
        self._press_attach = self._attach   # 记住挂靠状态（单击后恢复）
        self._detach()
        self._press = (e.x_root, e.y_root, time.time())
        self._moved = False
        self._drag_offset = (e.x_root - self.pos()[0],
                             e.y_root - self.pos()[1])

    def _on_drag(self, e):
        if self._press is None:
            return
        dx = e.x_root - self._press[0]
        dy = e.y_root - self._press[1]
        if abs(dx) > 6 or abs(dy) > 6:
            if not self._moved:
                # 开拖就把气泡关掉：大气泡窗口跟着高速移动会闪出黑框
                if self.dialog is not None and self.dialog.top.winfo_exists():
                    try:
                        self.dialog.close()
                    except Exception:
                        pass
                self.set_state("drag")   # 开始移动才切换拖动图（窗口随之放大）
                # 窗口放大后重新锚定抓取点，人物不跳位
                self._drag_offset = (e.x_root - self.pos()[0],
                                     e.y_root - self.pos()[1])
            self._moved = True
        x = e.x_root - self._drag_offset[0]
        y = e.y_root - self._drag_offset[1]
        # 合并高频移动事件（快速拖动时每秒上百次），避免窗口来不及
        # 重绘闪出背景色（近黑）形成黑框；攒着的位置由定时器补齐
        self._pending_pos = (x, y)
        now = time.time()
        if now - self._last_move_t >= 0.016:
            self._flush_drag_move()
        elif not self._move_job:
            delay = max(1, int((0.016 - (now - self._last_move_t)) * 1000))
            self._move_job = self.win.after(delay, self._flush_drag_move)

    def _flush_drag_move(self):
        self._move_job = None
        if self._pending_pos is None or self._press is None:
            return
        self._last_move_t = time.time()
        x, y = self._pending_pos
        self.move_to(x, y)

    def _on_release(self, e):
        if self._press is None:
            return
        # 先取消定时器并把攒着的最后一次移动落位，吸附/落点计算才准确
        if self._move_job:
            try:
                self.win.after_cancel(self._move_job)
            except Exception:
                pass
            self._move_job = None
        self._flush_drag_move()
        self._pending_pos = None
        dt = time.time() - self._press[2]
        self._press = None
        # 拖动放大的窗口先缩回常规大小（保持中心），再计算吸附/落点
        if self._cur_size != self._default_size():
            self._resize_to(self._default_size())
        if not self._moved and dt < 0.6:
            # 单击：人物抖一下 + 弹出气泡
            self.set_state("idle")
            self.open_dialog(advance=True)
            # 单击后恢复挂靠（如果之前挂在 Word 窗口上）
            if self._press_attach:
                hwnd, edge = self._press_attach
                if _user32.IsWindow(hwnd):
                    self._attach_to(hwnd, edge)
            self._press_attach = None
            self._shake()
        else:
            self._snap_or_fall()

    def _snap_or_fall(self):
        x, y = self.pos()
        w, h = self.size
        OVL = self.WORD_EDGE_OVERLAP
        # 候选：(距离, 目标x, 目标y, 挂靠信息)——挂靠信息非空表示挂 Word 窗口
        candidates = []

        # Word 窗口上边缘（仅未最大化的窗口；挂上后跟随窗口移动）
        for hwnd, (l, t, r, b) in find_word_windows():
            if l < -10000 or is_maximized(hwnd):   # 最小化 / 最大化：不可挂
                continue
            # 水平方向：与窗口有交叠即可（左右各放宽 60px，窗口外侧一点也能挂）
            h_overlap = (x + w > l - 60) and (x < r + 60)
            if h_overlap:
                if y + h <= t:
                    # 桌宠整个在窗口上方（窗口外）：
                    # 按“桌宠底边到窗口顶边”的空隙算，SNAP_ABOVE 内都能吸附；
                    # 空隙按比例折算成候选距离（空隙越大优先级越低）
                    gap = t - (y + h)
                    dist = gap * self.SNAP_DIST / self.SNAP_ABOVE
                else:
                    # 桌宠压着 / 低于窗口顶边：按桌宠顶边到窗口顶边算
                    dist = abs(y - t)
                candidates.append((dist, x, t - h + OVL, (hwnd, "top")))

        # 屏幕底部（固定落位：任务栏处稍微沉入一点）
        # 注意：必须用真实距离（abs），人物在目标线上方时距离不能算 0，
        # 否则底部永远赢得吸附竞争，窗口就吸不上了
        bottom_dist = abs(y - self.bottom_y())
        candidates.append((bottom_dist, x, self.bottom_y(), None))

        dist, tx, ty, attach = min(candidates, key=lambda c: c[0])
        if dist <= self.SNAP_DIST or bottom_dist <= self.BOTTOM_SNAP_DIST:
            # 只有挂到 Word 窗口边缘才播 hang 动画；
            # 停在屏幕底部属于普通落位（idle 动画）
            self._slide_to(tx, ty, hang=attach is not None, attach=attach)
        else:
            # 松手掉落（无回弹，落地即停）
            self._start_fall()

    def _slide_to(self, tx, ty, hang=False, attach=None):
        self.cancel_anim()
        token = self._anim_token
        x0, y0 = self.pos()
        steps = 8

        def step(i):
            if token != self._anim_token or not self.visible:
                return
            if i > steps:
                if hang:
                    self.set_state("hang")
                elif not attach:
                    self.set_state("idle")   # 停在屏幕底部 → 待机动画
                if attach:
                    self._attach_to(*attach)
                return
            t = i / steps
            x = int(x0 + (tx - x0) * t)
            y = int(y0 + (ty - y0) * t)
            self.move_to(x, y)
            self.win.after(15, step, i + 1)

        step(0)

    def _shake(self):
        """单击：人物向上轻轻拉伸一下（单次、小幅）。
        窗口大小全程不变（拉伸图画在固定尺寸画布上、底部对齐），
        避免透明窗口缩放时背景色闪现形成黑框。"""
        pils = self.player.pil_frames
        if not pils:
            return
        self.cancel_anim()
        token = self._anim_token
        base = pils[(self.player.idx - 1) % len(pils)][0]   # 元组：(Image, dur)
        w0, h0 = base.size
        # 拉伸进度：平滑地起来一次再回落
        phases = [0.0, 0.45, 0.85, 1.0, 0.6, 0.0]

        def sized(p):
            # 高度小幅增加（向上拉伸），宽度略缩（变形感）
            return (max(1, int(w0 * (1 - 0.04 * p))),
                    max(1, int(h0 * (1 + 0.10 * p))))

        # 用带 alpha 的帧缩放 → defringe → 压平 → 贴回固定尺寸画布（底部居中）
        self._shake_photos = []
        for p in phases:
            im = base.resize(sized(p), Image.LANCZOS)
            im = _defringe(im)
            im = _flatten_transparent(im)
            canvas = Image.new("RGB", (w0, h0), PET_TRANS)
            canvas.paste(im, ((w0 - im.width) // 2, h0 - im.height))
            self._shake_photos.append(ImageTk.PhotoImage(canvas))

        self.player.pause()

        def step(i):
            if token != self._anim_token or not self.visible:
                self.player.resume()
                return
            if i >= len(phases):
                self.player.resume()
                return
            self.label.config(image=self._shake_photos[i])
            self.win.after(40, step, i + 1)

        step(0)

    # ---- 挂靠 Word 窗口 --------------------------------------------------
    def _attach_to(self, hwnd, edge):
        """把桌宠挂到 Word 窗口 hwnd 的 edge 边缘上，并跟随窗口移动。"""
        if is_maximized(hwnd):
            return
        rect = window_rect(hwnd)
        if rect is None:
            return
        l, t, r, b = rect
        if l < -10000:
            return
        w, h = self.size
        if edge in ("left", "right"):
            self._attach_offset = max(0, min(self.pos()[1] - t, (b - t) - h))
        else:
            self._attach_offset = max(0, min(self.pos()[0] - l, (r - l) - w))
        self._attach = (hwnd, edge)
        self.set_state("hang")
        self._follow_word()

    def _detach(self):
        """取消挂靠（拖动 / 隐藏时调用）。"""
        self._attach = None
        if self._follow_job:
            try:
                self.win.after_cancel(self._follow_job)
            except Exception:
                pass
            self._follow_job = None

    def _follow_word(self):
        """挂靠期间循环：跟随 Word 窗口移动/缩放；窗口关闭或最小化则脱落。"""
        self._follow_job = None
        if not self._attach or not self.visible:
            return
        hwnd, edge = self._attach
        if not _user32.IsWindow(hwnd) or not _user32.IsWindowVisible(hwnd):
            self._detach()
            self.set_state("idle")
            return
        rect = window_rect(hwnd)
        if rect is None:
            self._detach()
            self.set_state("idle")
            return
        l, t, r, b = rect
        if l < -10000:
            # 最小化：脱落并回到屏幕下方中间
            self._detach()
            self.set_state("idle")
            x = (self.screen_w - self.size[0]) // 2
            self._slide_to(x, self.bottom_y(), hang=False)
            return
        if (r - l) < 120 or is_maximized(hwnd):   # 过窄 / 最大化：原地脱落
            self._detach()
            self.set_state("idle")
            return
        w, h = self.size
        OVL = self.WORD_EDGE_OVERLAP
        o = self._attach_offset
        if edge == "left":
            tx, ty = l - w + OVL, t + o
        elif edge == "right":
            tx, ty = r - OVL, t + o
        elif edge == "top":
            tx, ty = l + o, t - h + OVL
        else:
            tx, ty = l + o, b - OVL
        tx, ty = self.clamp_pos(tx, ty)
        if (tx, ty) != self.pos():
            self.move_to(tx, ty)
        self._follow_job = self.win.after(self.FOLLOW_INTERVAL, self._follow_word)

    def _start_fall(self):
        self.cancel_anim()
        token = self._anim_token
        self._vy = 0.0
        self.set_state("fall")
        self._fall_step(token)

    def _fall_step(self, token):
        if token != self._anim_token or not self.visible:
            return
        self._vy += 1.8
        x = self.pos()[0]
        y = self.win.winfo_y() + int(self._vy)
        bottom = self.bottom_y()
        if y >= bottom:
            # 落地即停（无回弹）
            self.move_to(x, bottom)
            self.set_state("idle")
            # fall 用的是放大窗口（203px），切回待机尺寸（154px）时
            # 按“保持中心”缩放会让底部上浮 ~25px，导致落点时高时低；
            # 这里按待机尺寸重新贴底，保证和吸附落位高度一致
            self.move_to(x, self.bottom_y())
            return
        self.move_to(x, y)
        self.win.after(16, self._fall_step, token)

    # ---- 右键菜单 --------------------------------------------------------
    def _on_menu(self, e):
        m = tk.Menu(self.root, tearoff=0)
        m.add_command(label="退出桌宠", command=self.root.destroy)
        try:
            m.tk_popup(e.x_root, e.y_root)
        finally:
            m.grab_release()

    # ---- 对话框 ----------------------------------------------------------
    def open_dialog(self, advance=True, auto_close_ms=3000):
        if self.dialog is not None and self.dialog.top.winfo_exists():
            self.dialog.top.deiconify()
            self.dialog.top.lift()
            self.dialog.top.focus_force()
            self.win.lift(self.dialog.top)   # 桌宠保持在气泡上层
        else:
            self.dialog = PetDialog(self, auto_close_ms)

    # ---- 闲置提醒：5 分钟没点桌宠，自动弹出气泡 1 秒 ----------------------
    def _check_idle(self):
        if (self.visible
                and (self.dialog is None
                     or not self.dialog.top.winfo_exists())
                and time.time() - self._last_click >= 300):
            self.open_dialog(auto_close_ms=1000)
        self.win.after(60000, self._check_idle)

    def on_dialog_closed(self, dlg):
        if self.dialog is dlg:
            self.dialog = None

    # ---- Word 打开 / 关闭 -------------------------------------------------
    def on_word_opened(self):
        if getattr(self, "_exit_job", None):
            # Word 又开了：取消退出计划
            try:
                self.root.after_cancel(self._exit_job)
            except Exception:
                pass
            self._exit_job = None
        self.show()

    def on_word_closed(self):
        # 需求：退出 Word 后桌宠自动退出。
        # 给 1 秒缓冲（防止 Word 只是重启/切换文档时被误关），
        # 期间 Word 重新打开则取消退出。
        self.hide()
        if getattr(self, "_exit_job", None):
            return
        self._exit_job = self.root.after(1000, self.root.destroy)


# ---------------------------------------------------------------------------
# 主程序
# ---------------------------------------------------------------------------
def main():
    # 高 DPI 下坐标才准确
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass

    cfg = load_config()

    root = tk.Tk()
    root.withdraw()
    root.title("Word 桌宠")

    events = queue.Queue()
    count_mode = cfg.get("counter", {}).get("mode", "chars")
    monitor = WordMonitor(events, count_mode)
    monitor.start()

    pet = PetWindow(root, cfg, monitor)

    # 后台预解码所有 GIF；主线程定时把解码结果转成 PhotoImage，
    # 之后状态切换（点击/拖动/挂边）全部零卡顿
    preload_gifs(cfg)

    def warm_photos():
        size = tuple(cfg["pet"].get("size", [140, 140]))
        for state, rel in cfg["pet"].get("gifs", {}).items():
            p = rel if os.path.isabs(rel) else os.path.join(BASE_DIR, rel)
            zoom, shift_y = _gif_params(cfg, state)
            if os.path.exists(p) and (p, size, zoom, shift_y) not in _GIF_PHOTOS:
                gif_photos(p, size, zoom, shift_y)
        root.after(2000, warm_photos)

    root.after(500, warm_photos)

    def check_events():
        try:
            while True:
                ev = events.get_nowait()
                if ev == "open":
                    pet.on_word_opened()
                elif ev == "close":
                    pet.on_word_closed()
        except queue.Empty:
            pass
        root.after(600, check_events)

    check_events()
    root.mainloop()


if __name__ == "__main__":
    main()
