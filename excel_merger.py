# -*- coding: utf-8 -*-
"""
Excel 合并工具 —— 扫描指定文件夹当前层级的 .xlsx，按表头一致性纵向合并。

运行环境：Windows / Python 3.9+
依赖：pandas, openpyxl
"""

from __future__ import annotations

import queue
import sys
import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog
from typing import Callable, List, Optional, Tuple


class _NullStream:
    """--noconsole 打包后 sys.stdout/stderr 为 None，任何写操作都会抛异常。"""

    def write(self, *_a, **_k):
        return 0

    def flush(self):
        pass

    def isatty(self):
        return False


if sys.stdout is None:
    sys.stdout = _NullStream()
if sys.stderr is None:
    sys.stderr = _NullStream()

import pandas as pd  # noqa: E402  （必须在流兜底之后导入）

# ---------------------------------------------------------------- 主题常量
BG = "#1b1d23"
PANEL = "#23262e"
PANEL_HI = "#2c3039"
FG = "#d7dbe0"
FG_DIM = "#8b93a1"
ACCENT = "#4c8dff"
OK = "#4ec9a0"
WARN = "#e5b567"
ERR = "#e06c75"
FONT_UI = ("Microsoft YaHei UI", 10)
FONT_MONO = ("Consolas", 9)

TEMP_PREFIX = "~$"
OUT_SHEET = "合并结果"


# ---------------------------------------------------------------- 核心逻辑
def scan_files(folder: Path) -> List[Path]:
    """只扫描当前层级（不递归），只取 .xlsx，忽略 ~$ 临时文件，按文件名升序。"""
    files = [
        p
        for p in folder.iterdir()
        if p.is_file()
        and p.suffix.lower() == ".xlsx"
        and not p.name.startswith(TEMP_PREFIX)
    ]
    files.sort(key=lambda p: p.name.lower())
    return files


def read_sheet(path: Path) -> pd.DataFrame:
    """读取第一个 Sheet，并删除全空行 / 全空列。"""
    df = pd.read_excel(path, sheet_name=0, engine="openpyxl")
    if df.shape[1] == 0:
        return df
    df = df.dropna(how="all")
    df = df.dropna(axis=1, how="all")
    return df.reset_index(drop=True)


def merge_frames(
    files: List[Path], log: Callable[[str, str], None]
) -> Tuple[Optional[pd.DataFrame], int, int]:
    """逐文件容错合并。返回 (合并结果, 成功数, 跳过数)。"""
    frames: List[pd.DataFrame] = []
    base_cols: Optional[List[str]] = None
    ok = 0
    skipped = 0

    for path in files:
        try:
            df = read_sheet(path)
        except Exception as exc:  # 文件损坏 / 被占用 / 非标准格式
            skipped += 1
            log(f"[跳过] {path.name} — 读取失败：{type(exc).__name__}: {exc}", "err")
            continue

        if df.empty:
            skipped += 1
            log(f"[跳过] {path.name} — 无有效数据（空表或全空行）", "warn")
            continue

        cols = list(df.columns)
        if base_cols is None:
            base_cols = cols
            log(f"标准表头已锁定：{len(cols)} 列（来自 {path.name}）", "ok")
        elif cols != base_cols:
            skipped += 1
            detail = "列数不同" if len(cols) != len(base_cols) else "列名或顺序不同"
            log(
                f"[跳过] {path.name} — 表头不一致（{detail}：本文件 {len(cols)} 列 / 标准 {len(base_cols)} 列）",
                "warn",
            )
            continue

        frames.append(df)
        ok += 1
        log(f"[合并] {path.name} — {len(df)} 行 × {len(cols)} 列", "ok")

    if not frames:
        return None, ok, skipped
    return pd.concat(frames, ignore_index=True, sort=False), ok, skipped


def _write(df: pd.DataFrame, target: Path) -> None:
    df.to_excel(target, index=False, sheet_name=OUT_SHEET)


def run_merge(folder: Path, log: Callable[[str, str], None]) -> None:
    """完整流程：扫描 → 合并 → 输出。所有异常均转为日志，不抛出。"""
    log(f"工作目录：{folder}", "info")

    try:
        files = scan_files(folder)
    except Exception as exc:
        log(f"目录扫描失败：{type(exc).__name__}: {exc}", "err")
        return

    if not files:
        log("未找到 .xlsx 文件（已忽略 ~$ 临时文件），任务结束。", "warn")
        return

    log(f"发现 {len(files)} 个 .xlsx 文件，开始读取…", "info")
    merged, ok, skipped = merge_frames(files, log)

    if merged is None:
        log(f"没有可合并的数据（成功 0，跳过 {skipped}）。", "err")
        return

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"合并结果_{stamp}.xlsx"
    parent = folder.parent if folder.parent != folder else folder

    saved: Optional[Path] = None
    target = parent / filename
    try:
        _write(merged, target)
        saved = target
    except Exception as exc:
        log(f"上级目录写入失败（{type(exc).__name__}），改用所选文件夹", "warn")
        target = folder / filename
        try:
            _write(merged, target)
            saved = target
        except Exception as exc2:
            log(f"写入失败，任务终止：{type(exc2).__name__}: {exc2}", "err")
            log(f"汇总：成功 {ok} 个，跳过 {skipped} 个（未生成文件）", "err")
            return

    log("", "info")
    log(f"汇总：成功 {ok} 个，跳过 {skipped} 个，合计 {len(merged)} 行。", "ok")
    log(f"输出：{saved}", "ok")


# ---------------------------------------------------------------- GUI
class App:
    POLL_MS = 80

    def __init__(self, root):
        self.root = root
        self.folder: Optional[Path] = None
        self.busy = False
        self.q: "queue.Queue[tuple]" = queue.Queue()

        self._build()
        self.root.after(self.POLL_MS, self._poll)

    # ---------------- 界面
    def _build(self):
        r = self.root
        r.title("Excel 合并工具")
        r.geometry("780x540")
        r.minsize(660, 460)
        r.configure(bg=BG)

        wrap = tk.Frame(r, bg=BG, padx=18, pady=16)
        wrap.pack(fill="both", expand=True)

        # 标题
        tk.Label(
            wrap,
            text="Excel 批量合并",
            bg=BG,
            fg=FG,
            font=("Microsoft YaHei UI", 15, "bold"),
            anchor="w",
        ).pack(fill="x")
        tk.Label(
            wrap,
            text="同一文件夹内表头一致的 .xlsx 纵向合并（仅当前层级，忽略 ~$ 临时文件）",
            bg=BG,
            fg=FG_DIM,
            font=("Microsoft YaHei UI", 9),
            anchor="w",
        ).pack(fill="x", pady=(2, 14))

        # 路径卡片
        card = tk.Frame(wrap, bg=PANEL, padx=12, pady=10)
        card.pack(fill="x")
        tk.Label(
            card, text="源文件夹", bg=PANEL, fg=FG_DIM, font=("Microsoft YaHei UI", 9), anchor="w"
        ).pack(fill="x")
        self.path_var = tk.StringVar(value="未选择")
        self.path_label = tk.Label(
            card,
            textvariable=self.path_var,
            bg=PANEL,
            fg=FG_DIM,
            font=FONT_MONO,
            anchor="w",
            justify="left",
        )
        self.path_label.pack(fill="x", pady=(3, 0))

        # 按钮行
        bar = tk.Frame(wrap, bg=BG)
        bar.pack(fill="x", pady=(12, 12))

        self.btn_pick = self._button(bar, "选择文件夹", self._pick, primary=False)
        self.btn_pick.pack(side="left")

        self.btn_start = self._button(bar, "开始合并", self._start, primary=True)
        self.btn_start.pack(side="left", padx=(10, 0))
        self.btn_start.config(state="disabled", bg=PANEL_HI, fg=FG_DIM)

        # 日志区
        logwrap = tk.Frame(wrap, bg=PANEL, padx=1, pady=1)
        logwrap.pack(fill="both", expand=True)

        self.log_box = tk.Text(
            logwrap,
            bg="#15171c",
            fg=FG,
            insertbackground=FG,
            font=FONT_MONO,
            relief="flat",
            wrap="word",
            padx=10,
            pady=8,
            state="disabled",
            highlightthickness=0,
        )
        sb = tk.Scrollbar(logwrap, command=self.log_box.yview, relief="flat", width=10)
        self.log_box.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.log_box.pack(side="left", fill="both", expand=True)

        self.log_box.tag_configure("time", foreground="#5f6673")
        self.log_box.tag_configure("info", foreground=FG)
        self.log_box.tag_configure("ok", foreground=OK)
        self.log_box.tag_configure("warn", foreground=WARN)
        self.log_box.tag_configure("err", foreground=ERR)

        self.log("就绪。请选择包含 .xlsx 的文件夹。", "info")

    def _button(self, parent, text, cmd, primary: bool):
        return tk.Button(
            parent,
            text=text,
            command=cmd,
            bg=ACCENT if primary else PANEL,
            fg="#ffffff" if primary else FG,
            activebackground="#3f7ae0" if primary else PANEL_HI,
            activeforeground="#ffffff" if primary else FG,
            font=FONT_UI,
            relief="flat",
            bd=0,
            padx=18,
            pady=7,
            cursor="hand2",
            disabledforeground=FG_DIM,
        )

    # ---------------- 日志（线程安全）
    def log(self, msg: str, level: str = "info"):
        self.q.put(("log", level, msg))

    def _append(self, level: str, msg: str):
        box = self.log_box
        box.configure(state="normal")
        box.insert("end", datetime.now().strftime("%H:%M:%S  "), "time")
        box.insert("end", msg + "\n", level)
        box.configure(state="disabled")
        box.see("end")

    def _poll(self):
        try:
            while True:
                item = self.q.get_nowait()
                if item[0] == "log":
                    self._append(item[1], item[2])
                elif item[0] == "done":
                    self._set_busy(False)
        except queue.Empty:
            pass
        self.root.after(self.POLL_MS, self._poll)

    # ---------------- 交互
    def _set_busy(self, busy: bool):
        self.busy = busy
        if busy:
            self.btn_start.config(text="合并中…", state="disabled", bg=PANEL_HI, fg=FG_DIM)
            self.btn_pick.config(state="disabled")
        else:
            self.btn_start.config(text="开始合并", state="normal", bg=ACCENT, fg="#ffffff")
            self.btn_pick.config(state="normal")

    def _pick(self):
        path = filedialog.askdirectory(title="选择包含 xlsx 文件的文件夹")
        if not path:
            return
        self.folder = Path(path)
        shown = str(self.folder)
        if len(shown) > 74:
            shown = shown[:26] + " … " + shown[-44:]
        self.path_var.set(shown)
        self.path_label.config(fg=FG)
        try:
            n = len(scan_files(self.folder))
            self.log(f"已选择：{self.folder}（检测到 {n} 个 .xlsx）", "info")
        except Exception as exc:
            self.log(f"已选择：{self.folder}（扫描失败：{exc}）", "warn")
        if not self.busy:
            self.btn_start.config(state="normal", bg=ACCENT, fg="#ffffff")

    def _start(self):
        if self.busy:
            return
        if self.folder is None:
            self.log("请先选择文件夹。", "warn")
            return
        self._set_busy(True)
        threading.Thread(target=self._worker, daemon=True).start()

    def _worker(self):
        try:
            run_merge(self.folder, self.log)
        except Exception as exc:  # 兜底，防止线程静默死亡
            import traceback

            self.log(f"未预期异常：{type(exc).__name__}: {exc}", "err")
            self.log(traceback.format_exc().strip().splitlines()[-1], "err")
        finally:
            self.q.put(("done", None, None))


# ---------------------------------------------------------------- 入口
def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
