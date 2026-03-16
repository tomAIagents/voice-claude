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
import struct
import wave
import tempfile
import threading
import subprocess
import queue
import tkinter as tk
import pyaudio
import keyboard
from faster_whisper import WhisperModel

# ─── Настройки ────────────────────────────────────────────────────────────────
HOTKEY      = 'caps lock'
MODEL_SIZE  = 'large-v3'
LANGUAGE    = 'ru'
SAMPLE_RATE = 16000
CHANNELS    = 1
CHUNK       = 1024

SCRIPT_DIR = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
AHK_SCRIPT = os.path.join(SCRIPT_DIR, "paste.ahk")

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
# ──────────────────────────────────────────────────────────────────────────────

model        = None
recording    = False
frames       = []
audio_thread = None
lock         = threading.Lock()
ui_queue     = queue.Queue()


def beep(freq=880, duration=100):
    try:
        import winsound
        winsound.Beep(freq, duration)
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
    tmp_fd, tmp_path = tempfile.mkstemp(suffix=".wav")
    os.close(tmp_fd)
    try:
        pa = pyaudio.PyAudio()
        wf = wave.open(tmp_path, "wb")
        wf.setnchannels(CHANNELS)
        wf.setsampwidth(pa.get_sample_size(pyaudio.paInt16))
        wf.setframerate(SAMPLE_RATE)
        wf.writeframes(b"".join(frames))
        wf.close()
        pa.terminate()
        segments, _ = model.transcribe(
            tmp_path, language=LANGUAGE,
            beam_size=1, vad_filter=True, word_timestamps=False,
        )
        return " ".join(seg.text.strip() for seg in segments).strip()
    finally:
        try:
            os.unlink(tmp_path)
        except OSError:
            pass


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


# ─── UI ───────────────────────────────────────────────────────────────────────

BARS     = 48
BAR_W    = 8
BAR_GAP  = 2
CANVAS_H = 56
WIN_W    = BARS * (BAR_W + BAR_GAP) + 30
WIN_H    = 230
BG       = '#f0f0f0'
BAR_COLOR = '#4a90d9'
BAR_IDLE  = '#d0d0d0'
BAR_LOAD  = '#f0a030'


class VoiceWindow:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("VoiceClaude")
        self.root.attributes('-topmost', True)
        self.root.resizable(False, False)
        self.root.configure(bg=BG)
        self._center()
        self._loading = True
        self._load_frame = 0

        self.status_var = tk.StringVar(value="  Загрузка модели...")

        tk.Label(
            self.root, textvariable=self.status_var,
            font=('Segoe UI', 11), fg='#444444', bg=BG
        ).pack(pady=(14, 4))

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

        # Запуск загрузки модели в фоне
        threading.Thread(target=load_model, daemon=True).start()

    def _animate_loading(self):
        if not self._loading:
            return
        self._load_frame += 1
        n = BARS
        levels = []
        for i in range(n):
            val = 0.15 + 0.6 * abs(math.sin((i + self._load_frame * 0.4) * 0.25))
            levels.append(val * 0.5)
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
        if msg['type'] == 'ready':
            self._loading = False
            self.levels = [0.0] * BARS
            self._draw_bars()
            self.status_var.set("  Зажми Caps Lock чтобы говорить...")
            keyboard.on_press_key(HOTKEY, on_press, suppress=True)
            keyboard.on_release_key(HOTKEY, on_release, suppress=True)

        elif msg['type'] == 'load_error':
            self._loading = False
            self.status_var.set(f"  Ошибка загрузки: {msg['text']}")

        elif msg['type'] == 'recording':
            self.root.lift()
            self.status_var.set("  Запись...  (отпусти Caps Lock)")
            self.text_area.configure(state=tk.NORMAL, bg='#ffffff')
            self.text_area.delete('1.0', tk.END)
            self.text_area.configure(state=tk.DISABLED)
            self.levels = [0.0] * BARS

        elif msg['type'] == 'level':
            self.levels.pop(0)
            self.levels.append(msg['value'])
            self._draw_bars()

        elif msg['type'] == 'transcribing':
            self.status_var.set("  Распознавание...")
            self.levels = [v * 0.3 for v in self.levels]
            self._draw_bars()

        elif msg['type'] == 'result':
            self.status_var.set("  Готово — редактируй и нажми Enter")
            self.levels = [0.0] * BARS
            self._draw_bars()
            self.text_area.configure(state=tk.NORMAL, bg='#ffffff')
            self.text_area.delete('1.0', tk.END)
            if msg['text']:
                self.text_area.insert('1.0', msg['text'])
            self.text_area.focus_set()
            self.text_area.mark_set(tk.INSERT, tk.END)

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

    def _center(self):
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        x = (sw - WIN_W) // 2
        y = (sh - WIN_H) // 2
        self.root.geometry(f"{WIN_W}x{WIN_H}+{x}+{y}")

    def run(self):
        self.root.mainloop()


VoiceWindow().run()
