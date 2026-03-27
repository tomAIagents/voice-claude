"""
VoiceClaude — Голосовой ввод для Claude Code в VSCode
------------------------------------------------------
Зажми Caps Lock → говори → отпусти → появится окно с текстом
Enter — вставить в терминал VSCode
Escape — отменить
"""

import os
import sys
import math
import json
import struct
import ctypes
import ctypes.wintypes as _wt
import winreg
import threading
import subprocess
import queue
import argparse
import tkinter as tk
import numpy as np
import pyaudio
import keyboard
from faster_whisper import WhisperModel
from PIL import Image, ImageDraw
import pystray

# ─── Аргументы командной строки ───────────────────────────────────────────────
_ap = argparse.ArgumentParser(add_help=False)
_ap.add_argument('--instance', type=int, default=0)
_ap.add_argument('--config',   type=str, default=None)
_args, _ = _ap.parse_known_args()
INSTANCE_ID = _args.instance   # 0 = одиночный режим, 1+ = мульти-экземпляр

# Каждый экземпляр получает уникальную F-клавишу (F9, F10, F11, F12...)
_INSTANCE_HOTKEYS = ['f9', 'f10', 'f11', 'f12']
_DEFAULT_HOTKEY   = _INSTANCE_HOTKEYS[INSTANCE_ID] if INSTANCE_ID < len(_INSTANCE_HOTKEYS) else 'f9'

# ─── Конфиг ───────────────────────────────────────────────────────────────────
SCRIPT_DIR  = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
_cfg_file   = _args.config or "config.json"
CONFIG_PATH = os.path.join(SCRIPT_DIR, _cfg_file)

_DEFAULTS = {"model": "small", "language": "ru", "hotkey": _DEFAULT_HOTKEY}

def load_config():
    try:
        with open(CONFIG_PATH, encoding='utf-8') as f:
            return {**_DEFAULTS, **json.load(f)}
    except FileNotFoundError:
        # Если это файл экземпляра — автоматически создаём с уникальным хоткеем
        if _args.config:
            base = os.path.join(SCRIPT_DIR, 'config.json')
            try:
                with open(base, encoding='utf-8') as f:
                    data = json.load(f)
            except Exception:
                data = dict(_DEFAULTS)
            data['hotkey'] = _DEFAULT_HOTKEY  # уникальная клавиша для этого экземпляра
            try:
                with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
                    json.dump(data, f, ensure_ascii=False, indent=4)
            except Exception:
                pass
            return {**_DEFAULTS, **data}
        return dict(_DEFAULTS)
    except Exception:
        return dict(_DEFAULTS)

cfg = load_config()

HOTKEY      = cfg['hotkey']
MODEL_SIZE  = cfg['model']
LANGUAGE    = cfg['language']
SAMPLE_RATE = 16000
CHANNELS    = 1
CHUNK       = 1024

AHK_SCRIPT = os.path.join(SCRIPT_DIR, "paste.ahk")
VBS_PATH   = os.path.join(SCRIPT_DIR, "VoiceClaude.vbs")
APP_NAME   = f"VoiceClaude #{INSTANCE_ID}" if INSTANCE_ID else "VoiceClaude"

def find_ahk():
    candidates = [
        os.path.expanduser(r"~\AppData\Local\Programs\AutoHotkey\v2\AutoHotkey64.exe"),
        r"C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe",
        r"C:\Program Files (x86)\AutoHotkey\v2\AutoHotkey64.exe",
    ]
    for p in candidates:
        if os.path.exists(p):
            return p
    return "AutoHotkey64.exe"

AHK_EXE = find_ahk()

# ─── Автозагрузка ─────────────────────────────────────────────────────────────
_REG_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"

def is_autostart():
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, _REG_KEY)
        winreg.QueryValueEx(key, APP_NAME)
        winreg.CloseKey(key)
        return True
    except OSError:
        return False

def _launch_cmd():
    """Команда запуска для реестра автозагрузки (instance-aware)."""
    py = sys.executable
    # Предпочитаем pythonw.exe для скрытого запуска
    pythonw = py.replace('python.exe', 'pythonw.exe').replace('Python.exe', 'pythonw.exe')
    if not os.path.exists(pythonw):
        pythonw = py
    script = os.path.join(SCRIPT_DIR, 'main.py')
    cmd = f'"{pythonw}" "{script}"'
    if INSTANCE_ID:
        cmd += f' --instance {INSTANCE_ID}'
    if _args.config:
        cmd += f' --config "{_args.config}"'
    return cmd


def set_autostart(enable):
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, _REG_KEY, 0, winreg.KEY_SET_VALUE)
        if enable:
            winreg.SetValueEx(key, APP_NAME, 0, winreg.REG_SZ, _launch_cmd())
        else:
            try:
                winreg.DeleteValue(key, APP_NAME)
            except OSError:
                pass
        winreg.CloseKey(key)
    except OSError:
        pass

# ──────────────────────────────────────────────────────────────────────────────

model        = None
recording    = False
frames       = []
audio_thread = None
lock         = threading.Lock()
ui_queue     = queue.Queue()
tray_icon    = None


def beep(freq=880, duration=100):
    try:
        import winsound
        winsound.Beep(freq, duration)
    except Exception:
        pass


def pin_to_all_desktops(hwnd):
    """Закрепить окно на всех виртуальных рабочих столах Windows 10/11."""
    try:
        class GUID(ctypes.Structure):
            _fields_ = [
                ('Data1', _wt.DWORD), ('Data2', _wt.WORD),
                ('Data3', _wt.WORD),  ('Data4', ctypes.c_ubyte * 8),
            ]

        ole32 = ctypes.windll.ole32
        ole32.CoInitialize(None)

        def mkguid(s):
            g = GUID()
            ole32.CLSIDFromString(s, ctypes.byref(g))
            return g

        CLSID_SHELL = mkguid('{C2F03A33-21F5-47FA-B4BB-156362A2F239}')  # ImmersiveShell
        IID_SP      = mkguid('{6D5140C1-7436-11CE-8034-00AA006009FA}')  # IServiceProvider
        IID_VDPA    = mkguid('{4CE81583-1E4C-4632-A621-07A53543148F}')  # IVirtualDesktopPinnedApps

        psp = ctypes.c_void_p()
        hr = ole32.CoCreateInstance(ctypes.byref(CLSID_SHELL), None, 4,
                                    ctypes.byref(IID_SP), ctypes.byref(psp))
        if hr or not psp.value:
            return

        # IServiceProvider::QueryService — vtable index 3
        QS = ctypes.WINFUNCTYPE(
            ctypes.HRESULT,
            ctypes.c_void_p, ctypes.POINTER(GUID), ctypes.POINTER(GUID),
            ctypes.POINTER(ctypes.c_void_p),
        )
        vt = ctypes.cast(ctypes.cast(psp, ctypes.POINTER(ctypes.c_void_p))[0],
                         ctypes.POINTER(ctypes.c_void_p))
        qs = QS(vt[3])

        pvdpa = ctypes.c_void_p()
        hr = qs(psp, ctypes.byref(IID_VDPA), ctypes.byref(IID_VDPA), ctypes.byref(pvdpa))
        if hr or not pvdpa.value:
            return

        # IVirtualDesktopPinnedApps::PinWindow — vtable index 6
        PW = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p, ctypes.c_void_p)
        vt2 = ctypes.cast(ctypes.cast(pvdpa, ctypes.POINTER(ctypes.c_void_p))[0],
                          ctypes.POINTER(ctypes.c_void_p))
        pw = PW(vt2[6])
        pw(pvdpa, ctypes.c_void_p(hwnd))
    except Exception:
        pass


def rms(data):
    shorts = struct.unpack(f'{len(data) // 2}h', data)
    mean_sq = sum(s * s for s in shorts) / len(shorts)
    return math.sqrt(mean_sq) / 32768.0


def load_model():
    global model
    try:
        m = WhisperModel(MODEL_SIZE, device="cpu", compute_type="int8")
        model = m
        ui_queue.put({'type': 'ready'})
    except Exception as ex:
        ui_queue.put({'type': 'load_error', 'text': str(ex)})


def record_audio():
    global frames
    pa = pyaudio.PyAudio()
    stream = pa.open(
        format=pyaudio.paInt16, channels=CHANNELS,
        rate=SAMPLE_RATE, input=True, frames_per_buffer=CHUNK,
    )
    frames = []
    while recording:
        data = stream.read(CHUNK, exception_on_overflow=False)
        frames.append(data)
        level = min(rms(data) * 6, 1.0)
        ui_queue.put({'type': 'level', 'value': level})
    stream.stop_stream()
    stream.close()
    pa.terminate()


def transcribe():
    if not frames:
        return ""
    audio_bytes = b"".join(frames)
    audio_np = np.frombuffer(audio_bytes, dtype=np.int16).astype(np.float32) / 32768.0
    segments, _ = model.transcribe(
        audio_np, language=LANGUAGE,
        beam_size=1, vad_filter=True, word_timestamps=False,
    )
    return " ".join(seg.text.strip() for seg in segments).strip()


def on_press(e):
    global recording, audio_thread
    if model is None:
        return
    with lock:
        if recording:
            return
        recording = True
        audio_thread = threading.Thread(target=record_audio, daemon=True)
        audio_thread.start()
    beep(880, 100)
    ui_queue.put({'type': 'recording'})


def on_release(e):
    global recording, audio_thread
    if model is None:
        return
    with lock:
        if not recording:
            return
        recording = False
    beep(440, 100)

    def process():
        if audio_thread:
            audio_thread.join()
        ui_queue.put({'type': 'transcribing'})
        try:
            text = transcribe()
        except Exception as ex:
            text = f"[Ошибка: {ex}]"
        ui_queue.put({'type': 'result', 'text': text})

    threading.Thread(target=process, daemon=True).start()


# ─── Иконка трея ──────────────────────────────────────────────────────────────

def _make_icon(bg, mic='white'):
    img = Image.new('RGBA', (64, 64), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.ellipse([2, 2, 62, 62], fill=bg)
    d.rounded_rectangle([22, 12, 42, 40], radius=10, fill=mic)
    d.arc([16, 30, 48, 54], start=0, end=180, fill=mic, width=4)
    d.line([32, 52, 32, 60], fill=mic, width=4)
    d.line([24, 60, 40, 60], fill=mic, width=4)
    return img


ICON_LOADING   = _make_icon('#f0a030')
ICON_IDLE      = _make_icon('#4a90d9')
ICON_RECORDING = _make_icon('#e74c3c')


def _set_tray(img, title):
    if tray_icon:
        tray_icon.icon  = img
        tray_icon.title = title


def start_tray(on_quit):
    global tray_icon

    def quit_action(icon, item):
        icon.stop()
        on_quit()

    def toggle_autostart(icon, item):
        set_autostart(not is_autostart())

    tray_icon = pystray.Icon(
        APP_NAME,
        ICON_LOADING,
        "VoiceClaude — загрузка...",
        menu=pystray.Menu(
            pystray.MenuItem(
                "Автозагрузка",
                toggle_autostart,
                checked=lambda item: is_autostart(),
            ),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Выход", quit_action),
        ),
    )
    tray_icon.run()


# ─── UI ───────────────────────────────────────────────────────────────────────

BARS      = 48
BAR_W     = 8
BAR_GAP   = 2
CANVAS_H  = 56
WIN_W     = BARS * (BAR_W + BAR_GAP) + 30
WIN_H     = 230
BG        = '#f0f0f0'
BAR_COLOR = '#4a90d9'
BAR_IDLE  = '#d0d0d0'
BAR_LOAD  = '#f0a030'


class VoiceWindow:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title(APP_NAME)
        self.root.attributes('-topmost', True)
        self.root.resizable(False, False)
        self.root.configure(bg=BG)
        self._center()
        self._loading = True
        self._load_frame = 0

        self.status_var = tk.StringVar(value="  Загрузка модели...")

        top_frame = tk.Frame(self.root, bg=BG)
        top_frame.pack(fill=tk.X, padx=15, pady=(14, 4))

        tk.Label(
            top_frame, textvariable=self.status_var,
            font=('Segoe UI', 11), fg='#444444', bg=BG
        ).pack(side=tk.LEFT)

        tk.Button(
            top_frame, text="⚙", font=('Segoe UI', 12),
            bg=BG, fg='#888888', relief=tk.FLAT, cursor='hand2',
            activebackground=BG, activeforeground='#444444',
            command=self._open_settings,
        ).pack(side=tk.RIGHT)

        self.canvas = tk.Canvas(
            self.root, width=WIN_W - 30, height=CANVAS_H,
            bg=BG, highlightthickness=0,
        )
        self.canvas.pack(padx=15)
        self.levels = [0.0] * BARS
        self._draw_bars()

        self.text_area = tk.Text(
            self.root, font=('Segoe UI', 12), wrap=tk.WORD,
            height=3, padx=10, pady=6, relief=tk.FLAT,
            bg='#ffffff', fg='#1e1e1e', insertbackground='black',
            state=tk.DISABLED,
        )
        self.text_area.pack(fill=tk.BOTH, expand=True, padx=15, pady=(8, 0))

        tk.Label(
            self.root, text="Enter — вставить    Esc — отмена",
            font=('Segoe UI', 9), fg='#999999', bg=BG
        ).pack(pady=(6, 10))

        self.root.bind('<Return>', self.on_confirm)
        self.root.bind('<Escape>', self.on_cancel)
        self.text_area.bind('<Return>', self.on_confirm)
        self.text_area.bind('<Escape>', self.on_cancel)

        self.root.after(40, self.process_queue)
        self.root.after(80, self._animate_loading)

        threading.Thread(target=start_tray, args=(self.root.destroy,), daemon=True).start()
        threading.Thread(target=load_model, daemon=True).start()

    def _animate_loading(self):
        if not self._loading:
            return
        self._load_frame += 1
        levels = [
            (0.15 + 0.6 * abs(math.sin((i + self._load_frame * 0.4) * 0.25))) * 0.5
            for i in range(BARS)
        ]
        self._draw_bars_custom(levels, BAR_LOAD)
        self.root.after(80, self._animate_loading)

    def _draw_bars(self):
        self._draw_bars_custom(self.levels, BAR_COLOR)

    def _draw_bars_custom(self, levels, color):
        self.canvas.delete('all')
        ch = CANVAS_H
        for i, level in enumerate(levels):
            x = i * (BAR_W + BAR_GAP) + BAR_GAP
            bar_h = max(3, int(level * (ch - 6)))
            y_top = (ch - bar_h) // 2
            c = color if level > 0.02 else BAR_IDLE
            self.canvas.create_rectangle(
                x, y_top, x + BAR_W, y_top + bar_h,
                fill=c, outline='',
            )

    def process_queue(self):
        try:
            while True:
                msg = ui_queue.get_nowait()
                self.handle(msg)
        except queue.Empty:
            pass
        self.root.after(40, self.process_queue)

    def handle(self, msg):
        t = msg['type']

        if t == 'ready':
            self._loading = False
            self.levels = [0.0] * BARS
            self._draw_bars()
            self.status_var.set("  Зажми Caps Lock чтобы говорить...")
            _set_tray(ICON_IDLE, "VoiceClaude — готов")
            keyboard.on_press_key(HOTKEY, on_press, suppress=True)
            keyboard.on_release_key(HOTKEY, on_release, suppress=True)

        elif t == 'load_error':
            self._loading = False
            self.status_var.set(f"  Ошибка загрузки: {msg['text']}")

        elif t == 'recording':
            self.root.lift()
            self.status_var.set("  Запись...  (отпусти Caps Lock)")
            self.text_area.configure(state=tk.NORMAL, bg='#ffffff')
            self.text_area.delete('1.0', tk.END)
            self.text_area.configure(state=tk.DISABLED)
            self.levels = [0.0] * BARS
            _set_tray(ICON_RECORDING, "VoiceClaude — запись...")

        elif t == 'level':
            self.levels.pop(0)
            self.levels.append(msg['value'])
            self._draw_bars()

        elif t == 'transcribing':
            self.status_var.set("  Распознавание...")
            self.levels = [v * 0.3 for v in self.levels]
            self._draw_bars()

        elif t == 'result':
            self.status_var.set("  Готово — редактируй и нажми Enter")
            self.levels = [0.0] * BARS
            self._draw_bars()
            self.text_area.configure(state=tk.NORMAL, bg='#ffffff')
            self.text_area.delete('1.0', tk.END)
            if msg['text']:
                self.text_area.insert('1.0', msg['text'])
            self.text_area.focus_set()
            self.text_area.mark_set(tk.INSERT, tk.END)
            _set_tray(ICON_IDLE, "VoiceClaude — готов")

    def on_confirm(self, event=None):
        text = self.text_area.get('1.0', tk.END).strip()
        self.status_var.set("  Зажми Caps Lock чтобы говорить...")
        self.text_area.configure(state=tk.NORMAL)
        self.text_area.delete('1.0', tk.END)
        self.text_area.configure(state=tk.DISABLED)
        self.levels = [0.0] * BARS
        self._draw_bars()
        if text:
            threading.Thread(
                target=lambda: subprocess.Popen([AHK_EXE, AHK_SCRIPT, text]),
                daemon=True
            ).start()
        return 'break'

    def on_cancel(self, event=None):
        self.status_var.set("  Зажми Caps Lock чтобы говорить...")
        self.text_area.configure(state=tk.NORMAL)
        self.text_area.delete('1.0', tk.END)
        self.text_area.configure(state=tk.DISABLED)
        self.levels = [0.0] * BARS
        self._draw_bars()

    def _open_settings(self):
        win = tk.Toplevel(self.root)
        win.title("Настройки")
        win.attributes('-topmost', True)
        win.resizable(False, False)
        win.configure(bg=BG)
        win.grab_set()

        pad = {'padx': 15, 'pady': 6}

        # Модель
        tk.Label(win, text="Модель Whisper:", font=('Segoe UI', 10),
                 bg=BG, anchor='w').pack(fill=tk.X, **pad)
        model_var = tk.StringVar(value=cfg.get('model', 'small'))
        model_menu = tk.OptionMenu(win, model_var,
                                   'tiny', 'base', 'small', 'medium', 'large-v3')
        model_menu.configure(font=('Segoe UI', 10), bg='white', relief=tk.FLAT)
        model_menu.pack(fill=tk.X, padx=15)

        tk.Label(win, text="  tiny — быстро, качество ниже\n"
                           "  small — быстро, хорошее качество ✓\n"
                           "  medium / large-v3 — медленно, лучшее качество",
                 font=('Segoe UI', 8), fg='#888888', bg=BG, justify=tk.LEFT,
                 ).pack(fill=tk.X, padx=15, pady=(0, 6))

        # Язык
        tk.Label(win, text="Язык распознавания:", font=('Segoe UI', 10),
                 bg=BG, anchor='w').pack(fill=tk.X, **pad)
        lang_var = tk.StringVar(value=cfg.get('language', 'ru'))
        lang_menu = tk.OptionMenu(win, lang_var, 'ru', 'en', 'uk', 'de', 'fr', 'es', 'zh')
        lang_menu.configure(font=('Segoe UI', 10), bg='white', relief=tk.FLAT)
        lang_menu.pack(fill=tk.X, padx=15)

        # Хоткей
        tk.Label(win, text="Горячая клавиша:", font=('Segoe UI', 10),
                 bg=BG, anchor='w').pack(fill=tk.X, **pad)
        hotkey_var = tk.StringVar(value=cfg.get('hotkey', 'caps lock'))
        tk.Entry(win, textvariable=hotkey_var, font=('Segoe UI', 10),
                 relief=tk.FLAT, bg='white').pack(fill=tk.X, padx=15)

        # Разделитель
        tk.Frame(win, bg='#dddddd', height=1).pack(fill=tk.X, padx=15, pady=12)

        # Кнопка сохранить
        msg_var = tk.StringVar()
        tk.Label(win, textvariable=msg_var, font=('Segoe UI', 9),
                 fg='#27ae60', bg=BG).pack()

        def save():
            cfg['model']    = model_var.get()
            cfg['language'] = lang_var.get()
            cfg['hotkey']   = hotkey_var.get().strip()
            try:
                with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
                    json.dump(cfg, f, ensure_ascii=False, indent=4)
                msg_var.set("Сохранено. Перезапустите программу.")
            except Exception as ex:
                msg_var.set(f"Ошибка: {ex}")

        tk.Button(
            win, text="Сохранить", font=('Segoe UI', 10, 'bold'),
            bg='#4a90d9', fg='white', relief=tk.FLAT, cursor='hand2',
            activebackground='#357abd', activeforeground='white',
            padx=20, pady=6, command=save,
        ).pack(pady=(4, 15))

    def _center(self):
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        x = (sw - WIN_W) // 2
        y = (sh - WIN_H) // 2
        self.root.geometry(f"{WIN_W}x{WIN_H}+{x}+{y}")

    def run(self):
        self.root.mainloop()


VoiceWindow().run()
