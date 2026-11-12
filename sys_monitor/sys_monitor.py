# -*- coding: utf-8 -*-
"""
SysMonitor - 极客风桌面系统监控悬浮窗
依赖: tkinter (标准库) + psutil
打包: pyinstaller --noconfirm --clean --onefile --noconsole --name SysMonitor main.py
交互: 左键任意位置拖动 / 双击关闭(5px 位移防抖)
"""

import ctypes
import queue
import threading
import time
import tkinter as tk

import psutil

# ---------------- DPI 适配(必须在 Tk 创建前) ----------------
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)   # PER_MONITOR_AWARE_V2
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

# ---------------- 常量 ----------------
BG = "#000000"
FG = "#00FF00"
FONT = ("Consolas", 13)
WIDTH = 22            # 固定字符宽度,防抖动
ALPHA = 0.85
MARGIN = 20
INTERVAL = 1.0
DASH = "--"
MB = 1048576.0


# ---------------- 后台采集 ----------------
def worker(q, stop_evt):
    prev_net, prev_t = None, None

    # 首次调用建立 cpu_percent / 网络的基线
    try:
        psutil.cpu_percent(interval=None)
        prev_net = psutil.net_io_counters()
        prev_t = time.time()
    except Exception:
        prev_net, prev_t = None, None

    while not stop_evt.is_set():
        d = {"time": time.strftime("%Y-%m-%d %H:%M:%S")}

        try:
            d["cpu"] = float(psutil.cpu_percent(interval=None))
        except Exception:
            d["cpu"] = None

        try:
            f = psutil.cpu_freq()
            d["freq"] = float(f.current) if f else None      # MHz
        except Exception:
            d["freq"] = None

        try:
            vm = psutil.virtual_memory()
            d["mem_used"] = vm.used / 1073741824.0
            d["mem_total"] = vm.total / 1073741824.0
        except Exception:
            d["mem_used"] = d["mem_total"] = None

        now = time.time()
        try:
            cur = psutil.net_io_counters()
        except Exception:
            cur = None

        d["up"] = d["down"] = None
        if cur is not None and prev_net is not None:
            try:
                dt = now - prev_t
                if dt > 0:
                    d["up"] = max((cur.bytes_sent - prev_net.bytes_sent) / dt / MB, 0.0)
                    d["down"] = max((cur.bytes_recv - prev_net.bytes_recv) / dt / MB, 0.0)
            except Exception:
                pass
        if cur is not None:
            prev_net, prev_t = cur, now

        try:
            q.put_nowait(d)
        except Exception:
            pass

        stop_evt.wait(INTERVAL)


def num(v, width, prec):
    try:
        return ("%*.*f" % (width, prec, v)) if v is not None else DASH.rjust(width)
    except Exception:
        return DASH.rjust(width)


# ---------------- UI ----------------
class Monitor(tk.Tk):
    def __init__(self):
        tk.Tk.__init__(self)

        self.overrideredirect(True)
        self.configure(bg=BG)
        try:
            self.attributes("-alpha", ALPHA)
        except Exception:
            pass
        self.attributes("-topmost", True)

        # 跟随系统 DPI,避免字体发虚
        try:
            self.tk.call("tk", "scaling", self.winfo_fpixels("1i") / 72.0)
        except Exception:
            pass

        self.labels = []
        for _ in range(4):
            lb = tk.Label(
                self, text=" " * WIDTH, font=FONT, fg=FG, bg=BG,
                width=WIDTH, anchor="w", justify="left",
                padx=2, pady=0, bd=0, highlightthickness=0,
            )
            lb.pack(anchor="w", fill="x")
            self.labels.append(lb)

        self.update_idletasks()
        w, h = self.winfo_reqwidth(), self.winfo_reqheight()
        self.geometry("%dx%d%+d%+d" % (w, h, max(self.winfo_screenwidth() - w - MARGIN, 0), MARGIN))

        self._pressed = False
        self._moved = 0
        self._last_up = 0.0
        self._last_up_moved = 99
        self._bind_all(self)

        self.q = queue.Queue(maxsize=8)
        self.stop_evt = threading.Event()
        threading.Thread(target=worker, args=(self.q, self.stop_evt), daemon=True).start()

        self.protocol("WM_DELETE_WINDOW", self.close)
        self.after(120, self._pump)

    # ---- 拖动 / 双击 ----
    def _bind_all(self, widget):
        widget.bind("<ButtonPress-1>", self._on_press, add="+")
        widget.bind("<B1-Motion>", self._on_motion, add="+")
        widget.bind("<ButtonRelease-1>", self._on_release, add="+")
        for ch in widget.winfo_children():
            self._bind_all(ch)

    def _on_press(self, ev):
        self._ox, self._oy = ev.x_root, ev.y_root
        self._wx, self._wy = self.winfo_x(), self.winfo_y()
        self._moved = 0
        self._pressed = True

    def _on_motion(self, ev):
        if not self._pressed:
            return
        dx, dy = ev.x_root - self._ox, ev.y_root - self._oy
        self._moved = max(self._moved, abs(dx) + abs(dy))
        try:
            self.geometry("%+d%+d" % (max(self._wx + dx, 0), max(self._wy + dy, 0)))
        except Exception:
            pass

    def _on_release(self, ev):
        self._pressed = False
        moved = self._moved
        now = time.time()
        # 防抖:本轮与上一轮按压位移均 <5px,且间隔 <0.45s 才算双击
        if moved < 5 and self._last_up_moved < 5 and (now - self._last_up) < 0.45:
            self.close()
            return
        self._last_up = now
        self._last_up_moved = moved

    # ---- 刷新 ----
    def _pump(self):
        try:
            for _ in range(4):
                try:
                    self._render(self.q.get_nowait())
                except queue.Empty:
                    break
        except Exception:
            pass
        try:
            self.attributes("-topmost", True)
        except Exception:
            pass
        try:
            self.after(200, self._pump)
        except Exception:
            pass

    def _render(self, d):
        try:
            lines = [
                "CPU %s%%  %sGHz" % (num(d.get("cpu"), 5, 1),
                                      num(None if d.get("freq") is None else d["freq"] / 1000.0, 5, 2)),
                "MEM %s/%sGB" % (num(d.get("mem_used"), 5, 1), num(d.get("mem_total"), 5, 1)),
                "NET \u2191%s \u2193%s MB/s" % (num(d.get("up"), 5, 2), num(d.get("down"), 5, 2)),
                d.get("time") or time.strftime("%H:%M:%S"),
            ]
            for lb, s in zip(self.labels, lines):
                lb.configure(text=s.ljust(WIDTH)[:WIDTH])
        except Exception:
            for lb in self.labels:
                try:
                    lb.configure(text=DASH.ljust(WIDTH))
                except Exception:
                    pass

    def close(self):
        try:
            self.stop_evt.set()
        except Exception:
            pass
        try:
            self.destroy()
        except Exception:
            pass


if __name__ == "__main__":
    try:
        Monitor().mainloop()
    except Exception:
        pass
