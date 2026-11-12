# -*- coding: utf-8 -*-
"""
报名照片 / 证件照 一键处理助手
依赖：Python 3.8+ / Pillow
环境：Windows（tkinter 自带）
"""

import io
import os
import queue
import threading
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, ttk

from PIL import Image, ImageChops, ImageOps

APP_TITLE = "报名照 / 证件照 一键处理"

PRESETS = {
    "一寸 295x413": (295, 413),
    "二寸 413x579": (413, 579),
    "小一寸 260x378": (260, 378),
    "自定义": None,
}

BG_COLORS = {
    "不换": None,
    "红底": (255, 0, 0),
    "蓝底": (67, 142, 219),
    "白底": (255, 255, 255),
}

BG_TOLERANCE = 30      # 颜色容差（每通道最大差值）
BG_MIN_COV = 0.5       # 四周环带主色最低占比，低于此值判为背景不纯
Q_MIN, Q_MAX = 5, 95   # 二分 quality 区间
MAX_ROUNDS = 8         # 降分辨率重试上限，绝不死循环

# ---------- 深色配色 ----------
C_BG = "#1b1b1f"
C_PANEL = "#26262c"
C_FG = "#e8e8ec"
C_DIM = "#8b8b95"
C_ACCENT = "#2d6cdf"


# ============================ 图像核心 ============================

def encode_jpeg(img, quality, dpi=(300, 300)):
    """统一出口：任何保存前强制转 RGB，兼容 RGBA / P / LA。"""
    if img.mode != "RGB":
        img = img.convert("RGB")
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=quality, optimize=True, dpi=dpi)
    return buf.getvalue()


def crop_to_ratio(img, tw, th):
    """居中裁切到目标宽高比（避免直接拉伸导致人脸变形）。"""
    w, h = img.size
    target = tw / th
    cur = w / h
    if abs(cur - target) < 1e-6:
        return img
    if cur > target:
        nw = max(1, int(round(h * target)))
        left = (w - nw) // 2
        return img.crop((left, 0, left + nw, h))
    nh = max(1, int(round(w / target)))
    top = (h - nh) // 2
    return img.crop((0, top, w, top + nh))


def _corner_means(img, tol):
    """四角小块均值，返回 (means, 是否互相一致)。"""
    w, h = img.size
    s = max(2, min(w, h) // 40)
    if w < s * 2 or h < s * 2:
        return [], False
    boxes = [(0, 0), (w - s, 0), (0, h - s), (w - s, h - s)]
    means = []
    for x, y in boxes:
        patch = img.crop((x, y, x + s, y + s))
        px = list(patch.getdata())
        n = len(px)
        means.append(tuple(sum(p[i] for p in px) // n for i in range(3)))
    avg = tuple(sum(m[i] for m in means) // len(means) for i in range(3))
    agree = all(max(abs(m[i] - avg[i]) for i in range(3)) <= tol for m in means)
    return means, agree


def detect_bg(img, tol=BG_TOLERANCE, ring=0.04, min_cov=BG_MIN_COV):
    """
    判定背景主色：取画面四周环带的众数色为准，不再要求四角完全一致
    （证件照里肩膀/头发常压住下角，旧逻辑会误判为"背景不一致"直接跳过）。
    返回 (bg_rgb, coverage, corners_agree)；无法判定时 bg_rgb 为 None。
    """
    w, h = img.size
    sc = min(1.0, 240.0 / max(w, h))
    sm = img if sc >= 1.0 else img.resize(
        (max(8, int(w * sc)), max(8, int(h * sc))), Image.BILINEAR)
    sw, sh = sm.size
    b = max(1, int(round(min(sw, sh) * ring)))
    px = sm.load()

    ring_px = []
    for y in range(sh):
        edge_y = y < b or y >= sh - b
        for x in range(sw):
            if edge_y or x < b or x >= sw - b:
                ring_px.append(px[x, y])
    if not ring_px:
        return None, 0.0, False

    buckets = {}
    for p in ring_px:
        k = (p[0] >> 4, p[1] >> 4, p[2] >> 4)
        buckets[k] = buckets.get(k, 0) + 1
    key = max(buckets, key=buckets.get)
    cov = buckets[key] / len(ring_px)

    group = [p for p in ring_px
             if (p[0] >> 4, p[1] >> 4, p[2] >> 4) == key]
    n = len(group)
    bg = tuple(sum(p[i] for p in group) // n for i in range(3))

    _, agree = _corner_means(sm, tol)
    if agree or cov >= min_cov:
        return bg, cov, agree
    return None, cov, agree


def replace_bg(img, src_rgb, dst_rgb, tol=BG_TOLERANCE):
    """纯本地像素计算：容差内的像素替换为 dst_rgb，其余原样保留。返回 (图, 替换占比)。"""
    diff = ImageChops.difference(img, Image.new("RGB", img.size, src_rgb))
    r, g, b = diff.split()
    maxdiff = ImageChops.lighter(ImageChops.lighter(r, g), b)
    mask = maxdiff.point(lambda v: 255 if v <= tol else 0)
    out = img.copy()
    out.paste(Image.new("RGB", img.size, dst_rgb), (0, 0), mask)
    hit = mask.histogram()[255]
    total = img.width * img.height
    return out, (hit / total if total else 0.0)


def compress_to_target(img, target_bytes, log):
    """
    二分 quality(5~95) 逼近目标体积；q=5 仍超标则等比降分辨率重来，最多 MAX_ROUNDS 轮。
    返回 ((w, h), data, None) 成功；((0, 0), data, 实际最小字节) 失败。
    """
    scale = 1.0
    last = b""
    for rnd in range(1, MAX_ROUNDS + 1):
        w = max(1, int(round(img.width * scale)))
        h = max(1, int(round(img.height * scale)))
        cur = img if (w, h) == img.size else img.resize((w, h), Image.LANCZOS)

        base = encode_jpeg(cur, Q_MIN)
        last = base
        if len(base) <= target_bytes:
            lo, hi = Q_MIN, Q_MAX
            best, best_q = base, Q_MIN
            while lo <= hi:
                mid = (lo + hi) // 2
                data = encode_jpeg(cur, mid)
                if len(data) <= target_bytes:
                    best, best_q, lo = data, mid, mid + 1
                else:
                    hi = mid - 1
            log("  ✓ 第 %d 轮命中：%dx%d @ quality=%d → %.1f KB"
                % (rnd, w, h, best_q, len(best) / 1024))
            return (w, h), best, None

        ratio = (target_bytes / len(base)) ** 0.5 * 0.95
        ratio = min(max(ratio, 0.5), 0.95)
        log("  第 %d 轮 %dx%d @q%d = %.1f KB 仍超标 → 尺寸缩至 %.1f%%"
            % (rnd, w, h, Q_MIN, len(base) / 1024, ratio * 100))
        scale *= ratio

    return (0, 0), last, len(last)


def stamp_ms():
    now = datetime.now()
    return now.strftime("%Y%m%d_%H%M%S_") + "%03d" % (now.microsecond // 1000)


def short_path(p, n=46):
    return p if len(p) <= n else "…" + p[-n:]


# ============================ UI ============================

class App:
    def __init__(self, root):
        self.root = root
        self.q = queue.Queue()
        self.path = None
        self.busy = False

        self._init_style()
        self._build_ui()
        self.root.after(80, self._drain)

    # ---------- 样式 ----------
    def _init_style(self):
        self.root.configure(bg=C_BG)
        st = ttk.Style()
        try:
            st.theme_use("clam")
        except tk.TclError:
            pass
        st.configure(".", background=C_BG, foreground=C_FG,
                     fieldbackground=C_PANEL, bordercolor="#3a3a42",
                     lightcolor=C_PANEL, darkcolor=C_PANEL)
        st.configure("TFrame", background=C_BG)
        st.configure("TLabel", background=C_BG, foreground=C_FG)
        st.configure("Dim.TLabel", background=C_BG, foreground=C_DIM)
        st.configure("TButton", background=C_PANEL, foreground=C_FG,
                     borderwidth=1, focusthickness=0, padding=6)
        st.map("TButton",
               background=[("active", "#31313a"), ("disabled", "#202026")],
               foreground=[("disabled", "#585862")])
        st.configure("Accent.TButton", background=C_ACCENT, foreground="#ffffff",
                     borderwidth=0, padding=8)
        st.map("Accent.TButton",
               background=[("active", "#3a7cf0"), ("disabled", "#2a2a31")],
               foreground=[("disabled", "#5c5c66")])
        st.configure("TCombobox", fieldbackground=C_PANEL, background=C_PANEL,
                     foreground=C_FG, arrowcolor=C_FG, padding=4,
                     selectbackground=C_PANEL, selectforeground=C_FG)
        st.map("TCombobox",
               fieldbackground=[("readonly", C_PANEL), ("disabled", "#202026")],
               foreground=[("readonly", C_FG), ("disabled", "#585862")],
               selectbackground=[("readonly", C_PANEL)],
               selectforeground=[("readonly", C_FG)])
        st.configure("TEntry", fieldbackground=C_PANEL, foreground=C_FG,
                     insertcolor=C_FG, padding=4)
        st.map("TEntry",
               fieldbackground=[("disabled", "#202026")],
               foreground=[("disabled", "#585862")])
        st.configure("Vertical.TScrollbar", background=C_PANEL, troughcolor=C_BG,
                     bordercolor=C_BG, arrowcolor=C_DIM)

        self.root.option_add("*TCombobox*Listbox.background", C_PANEL)
        self.root.option_add("*TCombobox*Listbox.foreground", C_FG)
        self.root.option_add("*TCombobox*Listbox.selectBackground", C_ACCENT)
        self.root.option_add("*TCombobox*Listbox.selectForeground", "#ffffff")

    # ---------- 布局 ----------
    def _build_ui(self):
        self.root.title(APP_TITLE)
        self.root.geometry("560x640")
        self.root.minsize(500, 560)

        frm = ttk.Frame(self.root, padding=16)
        frm.pack(fill="both", expand=True)
        frm.columnconfigure(1, weight=1)
        frm.rowconfigure(8, weight=1)

        # 图片
        ttk.Label(frm, text="图片").grid(row=0, column=0, sticky="w", pady=5)
        self.var_path = tk.StringVar(value="未选择")
        ttk.Label(frm, textvariable=self.var_path, style="Dim.TLabel",
                  anchor="w").grid(row=0, column=1, sticky="we", pady=5)
        ttk.Button(frm, text="选择图片", command=self._pick).grid(
            row=1, column=0, columnspan=2, sticky="we", pady=(0, 12))

        # 尺寸规格
        ttk.Label(frm, text="尺寸规格").grid(row=2, column=0, sticky="w", pady=5)
        self.var_size = tk.StringVar(value="一寸 295x413")
        self.cb_size = ttk.Combobox(frm, textvariable=self.var_size,
                                    values=list(PRESETS.keys()),
                                    state="readonly", width=18)
        self.cb_size.grid(row=2, column=1, sticky="we", pady=5)

        # 宽 / 高
        ttk.Label(frm, text="宽 (px)").grid(row=3, column=0, sticky="w", pady=5)
        self.var_w = tk.StringVar(value="295")
        self.e_w = ttk.Entry(frm, textvariable=self.var_w, width=18)
        self.e_w.grid(row=3, column=1, sticky="we", pady=5)

        ttk.Label(frm, text="高 (px)").grid(row=4, column=0, sticky="w", pady=5)
        self.var_h = tk.StringVar(value="413")
        self.e_h = ttk.Entry(frm, textvariable=self.var_h, width=18)
        self.e_h.grid(row=4, column=1, sticky="we", pady=5)

        # 目标体积
        ttk.Label(frm, text="目标大小 (KB)").grid(row=5, column=0, sticky="w", pady=5)
        self.var_kb = tk.StringVar(value="20")
        ttk.Entry(frm, textvariable=self.var_kb, width=18).grid(
            row=5, column=1, sticky="we", pady=5)

        # 背景
        ttk.Label(frm, text="背景替换").grid(row=6, column=0, sticky="w", pady=5)
        self.var_bg = tk.StringVar(value="不换")
        ttk.Combobox(frm, textvariable=self.var_bg, values=list(BG_COLORS.keys()),
                     state="readonly", width=18).grid(
            row=6, column=1, sticky="we", pady=5)

        # 生成
        self.btn_go = ttk.Button(frm, text="一键生成", style="Accent.TButton",
                                 command=self._on_generate)
        self.btn_go.grid(row=7, column=0, columnspan=2, sticky="we", pady=(14, 12))

        # 日志
        logfrm = ttk.Frame(frm)
        logfrm.grid(row=8, column=0, columnspan=2, sticky="nsew")
        logfrm.rowconfigure(0, weight=1)
        logfrm.columnconfigure(0, weight=1)

        self.txt = tk.Text(logfrm, height=10, bg=C_PANEL, fg=C_FG,
                           insertbackground=C_FG, relief="flat", bd=0,
                           wrap="word", font=("Microsoft YaHei UI", 9),
                           padx=10, pady=8, state="disabled")
        self.txt.grid(row=0, column=0, sticky="nsew")
        sb = ttk.Scrollbar(logfrm, orient="vertical", command=self.txt.yview)
        sb.grid(row=0, column=1, sticky="ns")
        self.txt.configure(yscrollcommand=sb.set)

        self.var_size.trace_add("write", lambda *_: self._on_size_change())
        self._on_size_change()

    # ---------- 交互 ----------
    def _on_size_change(self):
        preset = PRESETS.get(self.var_size.get())
        if preset:
            self.var_w.set(str(preset[0]))
            self.var_h.set(str(preset[1]))
            self.e_w.configure(state="disabled")
            self.e_h.configure(state="disabled")
        else:
            self.e_w.configure(state="normal")
            self.e_h.configure(state="normal")

    def _pick(self):
        p = filedialog.askopenfilename(
            title="选择图片",
            filetypes=[("图片", "*.jpg *.jpeg *.png *.bmp *.webp *.tif *.tiff"),
                       ("所有文件", "*.*")])
        if p:
            self.path = p
            self.var_path.set(short_path(p))
            self._append("已选择：" + p)

    # ---------- 日志 / 队列 ----------
    def _log(self, msg):
        self.q.put(("log", msg))

    def _append(self, msg):
        self.txt.configure(state="normal")
        self.txt.insert("end", msg + "\n")
        self.txt.see("end")
        self.txt.configure(state="disabled")

    def _drain(self):
        try:
            while True:
                kind, payload = self.q.get_nowait()
                if kind == "log":
                    self._append(payload)
                elif kind == "done":
                    self.busy = False
                    self.btn_go.configure(state="normal")
                    if payload:
                        self._append("完成 → " + payload)
        except queue.Empty:
            pass
        self.root.after(80, self._drain)

    # ---------- 生成 ----------
    def _on_generate(self):
        if self.busy:
            return
        if not self.path or not os.path.isfile(self.path):
            self._append("⚠ 请先选择图片")
            return
        try:
            tw = self._to_int(self.var_w.get(), "宽度", 10, 5000)
            th = self._to_int(self.var_h.get(), "高度", 10, 5000)
            kb = self._to_kb(self.var_kb.get())
        except ValueError as e:
            self._append("⚠ " + str(e))
            return

        self.busy = True
        self.btn_go.configure(state="disabled")
        self._append("─" * 46)
        self._append("开始：%dx%d，目标 ≤%.1f KB，背景=%s"
                     % (tw, th, kb, self.var_bg.get())
                     if kb else "开始：%dx%d，不限制体积，背景=%s"
                     % (tw, th, self.var_bg.get()))

        threading.Thread(target=self._worker,
                         args=(self.path, tw, th, kb, self.var_bg.get()),
                         daemon=True).start()

    @staticmethod
    def _to_int(s, name, lo, hi):
        s = s.strip()
        if not s.isdigit():
            raise ValueError("%s 必须是正整数" % name)
        v = int(s)
        if not lo <= v <= hi:
            raise ValueError("%s 需在 %d~%d 之间" % (name, lo, hi))
        return v

    @staticmethod
    def _to_kb(s):
        s = s.strip()
        if not s:
            return 0
        try:
            v = float(s)
        except ValueError:
            raise ValueError("目标大小必须是数字（KB）")
        if v <= 0:
            raise ValueError("目标大小必须大于 0")
        return v

    # ---------- 工作线程 ----------
    def _worker(self, path, tw, th, kb, bg_name):
        try:
            self._log("▶ 载入：" + os.path.basename(path))
            with Image.open(path) as im:
                im = ImageOps.exif_transpose(im)
                img = im.convert("RGB")
            self._log("  原始 %dx%d，已转 RGB" % (img.width, img.height))

            img = crop_to_ratio(img, tw, th)
            self._log("  居中裁切至 %dx%d（目标比例 %d:%d）"
                      % (img.width, img.height, tw, th))

            dst = BG_COLORS.get(bg_name)
            if dst is not None:
                dom, cov, agree = detect_bg(img, BG_TOLERANCE)
                if dom is None:
                    self._log("警告：背景颜色不一致，未进行底色替换")
                    self._log("  原因：四周环带主色只占 %.0f%%（阈值 %.0f%%），四角也不一致。"
                              "可换一张纯色背景的照片，或手动修图"
                              % (cov * 100, BG_MIN_COV * 100))
                else:
                    img, ratio = replace_bg(img, dom, dst, BG_TOLERANCE)
                    self._log("  ✓ %s完成：基准 RGB%s，环带占比 %.0f%%，替换 %.1f%% 像素"
                              "（容差 %d）%s"
                              % (bg_name, dom, cov * 100, ratio * 100, BG_TOLERANCE,
                                 "" if agree else "；四角色差大，已按环带主色判定"))
                    if ratio > 0.60:
                        self._log("  ⚠ 替换比例偏高，浅色衣物可能被一并染色，请核对成片")

            img = img.resize((tw, th), Image.LANCZOS).convert("RGB")

            if kb > 0:
                target_bytes = int(kb * 1024)
                (w, h), data, err = compress_to_target(img, target_bytes, self._log)
                if err is not None:
                    self._log("  ✗ %d 轮后仍超标，放弃保存。实际最小 %.1f KB（目标 %.1f KB）"
                              % (MAX_ROUNDS, err / 1024, kb))
                    self.q.put(("done", None))
                    return
            else:
                (w, h), data = img.size, encode_jpeg(img, 92)
                self._log("  未限制体积，quality=92")

            out_dir = os.path.dirname(os.path.abspath(path)) or os.getcwd()
            out = os.path.join(out_dir, "报名照_%dx%d_%s.jpg" % (w, h, stamp_ms()))
            n = 1
            while os.path.exists(out):
                out = os.path.join(out_dir,
                                   "报名照_%dx%d_%s_%d.jpg" % (w, h, stamp_ms(), n))
                n += 1
            with open(out, "wb") as f:
                f.write(data)
            self._log("  ✓ 已保存：%.1f KB" % (len(data) / 1024))
            self.q.put(("done", out))
        except Exception as e:
            self._log("  ✗ 失败：%s: %s" % (type(e).__name__, e))
            self.q.put(("done", None))


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
