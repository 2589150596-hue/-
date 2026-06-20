#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Pomodoro Timer - 桌面番茄钟"""

import tkinter as tk
from tkinter import ttk, messagebox
import time
import threading
import winsound
import os


class PomodoroTimer:
    # 时间常量（秒）
    WORK_TIME = 25 * 60
    SHORT_BREAK = 5 * 60
    LONG_BREAK = 15 * 60

    def __init__(self, root):
        self.root = root
        self.root.title("番茄钟")
        self.root.geometry("400x540")
        self.root.resizable(False, False)
        self.root.configure(bg="#2c3e50")

        # 状态变量
        self.remaining = self.WORK_TIME
        self.running = False
        self.timer_thread = None
        self.current_mode = "work"  # work, short_break, long_break, custom
        self.completed_pomodoros = 0
        self.current_task = ""
        self.custom_time_seconds = 25 * 60

        self._build_ui()
        self._center_window()
        self._update_display()

    def _center_window(self):
        self.root.update_idletasks()
        w = self.root.winfo_width()
        h = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (w // 2)
        y = (self.root.winfo_screenheight() // 2) - (h // 2)
        self.root.geometry(f"{w}x{h}+{x}+{y}")

    def _build_ui(self):
        # 顶部模式标签
        self.mode_label = tk.Label(
            self.root,
            text="专注时间",
            font=("Microsoft YaHei", 20, "bold"),
            bg="#2c3e50",
            fg="#ecf0f1",
        )
        self.mode_label.pack(pady=(30, 10))

        # 任务输入框
        self.task_frame = tk.Frame(self.root, bg="#2c3e50")
        self.task_frame.pack(pady=10)
        tk.Label(
            self.task_frame,
            text="当前任务：",
            font=("Microsoft YaHei", 12),
            bg="#2c3e50",
            fg="#bdc3c7",
        ).pack(side=tk.LEFT)
        self.task_entry = tk.Entry(
            self.task_frame,
            font=("Microsoft YaHei", 12),
            width=20,
            justify="center",
            bg="#34495e",
            fg="#ecf0f1",
            insertbackground="#ecf0f1",
            relief=tk.FLAT,
        )
        self.task_entry.pack(side=tk.LEFT, padx=5)

        # 时间显示
        self.time_label = tk.Label(
            self.root,
            text="25:00",
            font=("Courier New", 64, "bold"),
            bg="#2c3e50",
            fg="#e74c3c",
        )
        self.time_label.pack(pady=20)

        # 进度条
        self.progress = ttk.Progressbar(
            self.root, length=300, mode="determinate", maximum=100
        )
        self.progress.pack(pady=10)
        self.style = ttk.Style()
        self.style.theme_use("clam")
        self.style.configure("red.Horizontal.TProgressbar", background="#e74c3c")
        self.style.configure("green.Horizontal.TProgressbar", background="#2ecc71")
        self.style.configure("blue.Horizontal.TProgressbar", background="#3498db")
        self.style.configure("purple.Horizontal.TProgressbar", background="#9b59b6")
        self.progress.configure(style="red.Horizontal.TProgressbar")

        # 按钮区域
        btn_frame = tk.Frame(self.root, bg="#2c3e50")
        btn_frame.pack(pady=20)

        self.start_btn = tk.Button(
            btn_frame,
            text="开始",
            font=("Microsoft YaHei", 14, "bold"),
            width=8,
            bg="#27ae60",
            fg="white",
            activebackground="#2ecc71",
            relief=tk.FLAT,
            cursor="hand2",
            command=self._toggle_timer,
        )
        self.start_btn.pack(side=tk.LEFT, padx=10)

        self.reset_btn = tk.Button(
            btn_frame,
            text="重置",
            font=("Microsoft YaHei", 14, "bold"),
            width=8,
            bg="#7f8c8d",
            fg="white",
            activebackground="#95a5a6",
            relief=tk.FLAT,
            cursor="hand2",
            command=self._reset_timer,
        )
        self.reset_btn.pack(side=tk.LEFT, padx=10)

        # 模式切换按钮
        mode_frame = tk.Frame(self.root, bg="#2c3e50")
        mode_frame.pack(pady=10)

        self.work_btn = tk.Button(
            mode_frame,
            text="专注",
            font=("Microsoft YaHei", 11),
            width=10,
            bg="#e74c3c",
            fg="white",
            relief=tk.FLAT,
            cursor="hand2",
            command=lambda: self._switch_mode("work"),
        )
        self.work_btn.pack(side=tk.LEFT, padx=5)

        self.short_btn = tk.Button(
            mode_frame,
            text="短休息",
            font=("Microsoft YaHei", 11),
            width=10,
            bg="#34495e",
            fg="white",
            relief=tk.FLAT,
            cursor="hand2",
            command=lambda: self._switch_mode("short_break"),
        )
        self.short_btn.pack(side=tk.LEFT, padx=5)

        self.long_btn = tk.Button(
            mode_frame,
            text="长休息",
            font=("Microsoft YaHei", 11),
            width=8,
            bg="#34495e",
            fg="white",
            relief=tk.FLAT,
            cursor="hand2",
            command=lambda: self._switch_mode("long_break"),
        )
        self.long_btn.pack(side=tk.LEFT, padx=5)

        self.custom_btn = tk.Button(
            mode_frame,
            text="自定义",
            font=("Microsoft YaHei", 11),
            width=8,
            bg="#34495e",
            fg="white",
            relief=tk.FLAT,
            cursor="hand2",
            command=lambda: self._switch_mode("custom"),
        )
        self.custom_btn.pack(side=tk.LEFT, padx=5)

        # 自定义时间设置
        custom_frame = tk.Frame(self.root, bg="#2c3e50")
        custom_frame.pack(pady=8)
        tk.Label(
            custom_frame,
            text="自定义时长：",
            font=("Microsoft YaHei", 11),
            bg="#2c3e50",
            fg="#bdc3c7",
        ).pack(side=tk.LEFT)
        self.custom_min_entry = tk.Entry(
            custom_frame,
            font=("Microsoft YaHei", 11),
            width=5,
            justify="center",
            bg="#34495e",
            fg="#ecf0f1",
            insertbackground="#ecf0f1",
            relief=tk.FLAT,
        )
        self.custom_min_entry.insert(0, "25")
        self.custom_min_entry.pack(side=tk.LEFT, padx=3)
        tk.Label(
            custom_frame,
            text="分钟",
            font=("Microsoft YaHei", 11),
            bg="#2c3e50",
            fg="#bdc3c7",
        ).pack(side=tk.LEFT)
        self.apply_custom_btn = tk.Button(
            custom_frame,
            text="应用",
            font=("Microsoft YaHei", 10),
            width=6,
            bg="#8e44ad",
            fg="white",
            activebackground="#9b59b6",
            relief=tk.FLAT,
            cursor="hand2",
            command=self._apply_custom_time,
        )
        self.apply_custom_btn.pack(side=tk.LEFT, padx=8)

        # 统计信息
        self.stats_label = tk.Label(
            self.root,
            text="今日完成：0 个番茄",
            font=("Microsoft YaHei", 12),
            bg="#2c3e50",
            fg="#bdc3c7",
        )
        self.stats_label.pack(pady=15)

    def _switch_mode(self, mode):
        if self.running:
            self._stop_timer()
        self.current_mode = mode
        if mode == "work":
            self.remaining = self.WORK_TIME
            self.mode_label.config(text="专注时间")
            self.time_label.config(fg="#e74c3c")
            self.progress.configure(style="red.Horizontal.TProgressbar")
            self._highlight_mode_btn(self.work_btn)
        elif mode == "short_break":
            self.remaining = self.SHORT_BREAK
            self.mode_label.config(text="短休息")
            self.time_label.config(fg="#2ecc71")
            self.progress.configure(style="green.Horizontal.TProgressbar")
            self._highlight_mode_btn(self.short_btn)
        elif mode == "long_break":
            self.remaining = self.LONG_BREAK
            self.mode_label.config(text="长休息")
            self.time_label.config(fg="#3498db")
            self.progress.configure(style="blue.Horizontal.TProgressbar")
            self._highlight_mode_btn(self.long_btn)
        elif mode == "custom":
            self.remaining = self.custom_time_seconds
            self.mode_label.config(text="自定义倒计时")
            self.time_label.config(fg="#9b59b6")
            self.progress.configure(style="purple.Horizontal.TProgressbar")
            self._highlight_mode_btn(self.custom_btn)
        self._update_display()

    def _highlight_mode_btn(self, active_btn):
        for btn, color in [
            (self.work_btn, "#e74c3c"),
            (self.short_btn, "#27ae60"),
            (self.long_btn, "#2980b9"),
            (self.custom_btn, "#8e44ad"),
        ]:
            if btn == active_btn:
                btn.config(bg=color)
            else:
                btn.config(bg="#34495e")

    def _toggle_timer(self):
        if self.running:
            self._stop_timer()
        else:
            self._start_timer()

    def _start_timer(self):
        self.running = True
        self.start_btn.config(text="暂停", bg="#f39c12", activebackground="#f1c40f")
        self.timer_thread = threading.Thread(target=self._timer_loop, daemon=True)
        self.timer_thread.start()

    def _stop_timer(self):
        self.running = False
        self.start_btn.config(text="开始", bg="#27ae60", activebackground="#2ecc71")

    def _reset_timer(self):
        self._stop_timer()
        if self.current_mode == "work":
            self.remaining = self.WORK_TIME
        elif self.current_mode == "short_break":
            self.remaining = self.SHORT_BREAK
        elif self.current_mode == "long_break":
            self.remaining = self.LONG_BREAK
        else:
            self.remaining = self.custom_time_seconds
        self._update_display()

    def _get_total_time(self):
        if self.current_mode == "work":
            return self.WORK_TIME
        elif self.current_mode == "short_break":
            return self.SHORT_BREAK
        elif self.current_mode == "long_break":
            return self.LONG_BREAK
        else:
            return self.custom_time_seconds

    def _timer_loop(self):
        total = self._get_total_time()
        while self.running and self.remaining > 0:
            time.sleep(1)
            if not self.running:
                return
            self.remaining -= 1
            self.root.after(0, self._update_display, total)
        if self.running and self.remaining <= 0:
            self.root.after(0, self._on_timer_complete)

    def _update_display(self, total=None):
        if total is None:
            total = self._get_total_time()
        minutes = self.remaining // 60
        seconds = self.remaining % 60
        self.time_label.config(text=f"{minutes:02d}:{seconds:02d}")
        progress_pct = ((total - self.remaining) / total) * 100 if total > 0 else 0
        self.progress["value"] = progress_pct
        # 更新窗口标题显示剩余时间
        self.root.title(f"[{minutes:02d}:{seconds:02d}] 番茄钟")

    def _on_timer_complete(self):
        self.running = False
        self.start_btn.config(text="开始", bg="#27ae60", activebackground="#2ecc71")
        self._play_beep()

        if self.current_mode == "work":
            self.completed_pomodoros += 1
            self.stats_label.config(text=f"今日完成：{self.completed_pomodoros} 个番茄")
            task = self.task_entry.get().strip()
            msg = f"恭喜完成一个番茄钟！" + (f"\n任务：{task}" if task else "")
            messagebox.showinfo("番茄钟完成", msg)
            # 自动建议休息
            if self.completed_pomodoros % 4 == 0:
                self._switch_mode("long_break")
            else:
                self._switch_mode("short_break")
        elif self.current_mode == "custom":
            messagebox.showinfo("倒计时结束", "自定义倒计时已结束！")
            self.remaining = self.custom_time_seconds
            self._update_display()
        else:
            messagebox.showinfo("休息结束", "休息结束，准备好开始新的专注了吗？")
            self._switch_mode("work")

    def _apply_custom_time(self):
        try:
            mins = int(self.custom_min_entry.get().strip())
            if mins <= 0 or mins > 180:
                messagebox.showwarning("输入错误", "请输入 1~180 之间的整数分钟")
                return
            self.custom_time_seconds = mins * 60
            if self.current_mode == "custom":
                self.remaining = self.custom_time_seconds
                self._update_display()
            else:
                self._switch_mode("custom")
        except ValueError:
            messagebox.showwarning("输入错误", "请输入有效的整数分钟数")

    def _play_beep(self):
        try:
            winsound.MessageBeep(winsound.MB_ICONEXCLAMATION)
            # 额外播放一个提示音
            for _ in range(3):
                winsound.Beep(800, 300)
                time.sleep(0.1)
        except Exception:
            pass


def main():
    root = tk.Tk()
    app = PomodoroTimer(root)
    root.mainloop()


if __name__ == "__main__":
    main()
