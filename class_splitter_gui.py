# -*- coding: utf-8 -*-
"""Small Tkinter front end for the class splitter."""

from __future__ import annotations

import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from class_splitter import (
    BALANCE_OPTIONS,
    ClassSplitError,
    create_template,
    split_csv_like,
    split_workbook,
)


class Tooltip:
    def __init__(self, widget: tk.Widget, text: str) -> None:
        self.widget = widget
        self.text = text
        self.tip: tk.Toplevel | None = None
        widget.bind("<Enter>", self.show)
        widget.bind("<Leave>", self.hide)

    def show(self, _event=None) -> None:
        if self.tip is not None:
            return
        x = self.widget.winfo_rootx() + 18
        y = self.widget.winfo_rooty() + 22
        self.tip = tk.Toplevel(self.widget)
        self.tip.wm_overrideredirect(True)
        self.tip.wm_geometry(f"+{x}+{y}")
        label = tk.Label(
            self.tip,
            text=self.text,
            justify=tk.LEFT,
            background="#FFFFE0",
            relief=tk.SOLID,
            borderwidth=1,
            padx=8,
            pady=6,
        )
        label.pack()

    def hide(self, _event=None) -> None:
        if self.tip is not None:
            self.tip.destroy()
            self.tip = None


class SplitterApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("均衡随机分班工具")
        self.geometry("720x380")
        self.minsize(680, 340)

        self.input_var = tk.StringVar()
        self.class_count_var = tk.StringVar(value="4")
        self.class_names_var = tk.StringVar(value="1班，2班，3班，4班")
        self.balance_var = tk.StringVar(value="ABC")
        self.seed_var = tk.StringVar()
        self.capacity_var = tk.StringVar()
        self.is_running = False

        self._build_ui()
        self.write_log("请手动选择名单文件；程序不会自动扫描本机已有 Excel 文件。")

    def _build_ui(self) -> None:
        frame = ttk.Frame(self, padding=16)
        frame.pack(fill=tk.BOTH, expand=True)
        frame.columnconfigure(1, weight=1)

        self._row_file(frame, 0, "名单文件", self.input_var, self.choose_input)

        ttk.Label(frame, text="班级数量").grid(row=1, column=0, sticky="w", pady=8)
        count_entry = ttk.Entry(frame, textvariable=self.class_count_var, width=12)
        count_entry.grid(row=1, column=1, sticky="w", pady=8)
        count_entry.bind("<FocusOut>", lambda _event: self.refresh_class_names())

        class_names_label_frame = ttk.Frame(frame)
        class_names_label_frame.grid(row=2, column=0, sticky="w", pady=8)
        ttk.Label(class_names_label_frame, text="班级名称").pack(side=tk.LEFT)
        class_names_help = ttk.Label(class_names_label_frame, text="?", foreground="#1F4E78", cursor="question_arrow")
        class_names_help.pack(side=tk.LEFT, padx=(6, 0))
        Tooltip(
            class_names_help,
            "班级名称请用中文逗号“，”隔开。\n\n"
            "例如：一班，二班，三班，四班，五班\n\n"
            "班级数量和班级名称个数必须一致。\n"
            "如果班级数量填 5，班级名称也要写 5 个。\n"
            "否则点击开始分班时会弹窗提醒先补齐或删减。",
        )
        ttk.Entry(frame, textvariable=self.class_names_var).grid(row=2, column=1, columnspan=2, sticky="ew", pady=8)

        ttk.Label(frame, text="根据哪些项分班").grid(row=3, column=0, sticky="w", pady=8)
        ttk.Entry(frame, textvariable=self.balance_var, width=18).grid(row=3, column=1, sticky="w", pady=8)
        ttk.Label(
            frame,
            text="A=性别  B=户籍大类  C=户籍细类  D=儿童人员类型",
            foreground="#555555",
        ).grid(row=3, column=1, sticky="w", padx=(150, 0), pady=8)

        seed_label_frame = ttk.Frame(frame)
        seed_label_frame.grid(row=4, column=0, sticky="w", pady=8)
        ttk.Label(seed_label_frame, text="随机种子").pack(side=tk.LEFT)
        seed_help = ttk.Label(seed_label_frame, text="?", foreground="#1F4E78", cursor="question_arrow")
        seed_help.pack(side=tk.LEFT, padx=(6, 0))
        Tooltip(
            seed_help,
            "随机种子用于复现分班结果。\n\n"
            "不填：每次分班可能不同。\n"
            "填写同一个数字：名单和设置不变时，结果相同。\n"
            "更换数字：得到另一套随机结果。",
        )
        ttk.Entry(frame, textvariable=self.seed_var, width=18).grid(row=4, column=1, sticky="w", pady=8)

        capacity_label_frame = ttk.Frame(frame)
        capacity_label_frame.grid(row=5, column=0, sticky="w", pady=8)
        ttk.Label(capacity_label_frame, text="每班上限").pack(side=tk.LEFT)
        capacity_help = ttk.Label(capacity_label_frame, text="?", foreground="#1F4E78", cursor="question_arrow")
        capacity_help.pack(side=tk.LEFT, padx=(6, 0))
        Tooltip(
            capacity_help,
            "每班上限用于限制单个班级最多人数。\n\n"
            "留空：不限制每班人数。\n"
            "填写数字：每班人数不会超过这个数字。\n"
            "如果 班级数量 × 每班上限 不够容纳学生，\n"
            "系统会自动新增班级，直到容量足够。\n\n"
            "例如：4个班、每班30人，若有125人参与分班，\n"
            "会自动新增5班。",
        )
        ttk.Entry(frame, textvariable=self.capacity_var, width=18).grid(row=5, column=1, sticky="w", pady=8)

        actions = ttk.Frame(frame)
        actions.grid(row=6, column=0, columnspan=3, sticky="ew", pady=(14, 8))
        ttk.Button(actions, text="生成模板", command=self.create_template_file).pack(side=tk.LEFT)
        self.start_button = ttk.Button(actions, text="开始分班", command=self.start_split)
        self.start_button.pack(side=tk.LEFT, padx=10)

        self.log = tk.Text(frame, height=8, wrap="word")
        self.log.grid(row=7, column=0, columnspan=3, sticky="nsew", pady=(8, 0))
        frame.rowconfigure(7, weight=1)

    def _row_file(self, parent: ttk.Frame, row: int, label: str, variable: tk.StringVar, command) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=8)
        ttk.Entry(parent, textvariable=variable).grid(row=row, column=1, sticky="ew", pady=8)
        ttk.Button(parent, text="选择", command=command).grid(row=row, column=2, sticky="e", padx=(8, 0), pady=8)

    def write_log(self, text: str) -> None:
        self.log.insert(tk.END, text + "\n")
        self.log.see(tk.END)

    def set_running(self, running: bool) -> None:
        self.is_running = running
        self.start_button.configure(state=tk.DISABLED if running else tk.NORMAL)

    def warn_large_file_if_needed(self, input_path: str) -> bool:
        path = Path(input_path)
        try:
            size_mb = path.stat().st_size / (1024 * 1024)
        except OSError:
            return True
        estimated_mb = size_mb * 35
        if size_mb < 15 and estimated_mb < 700:
            return True
        message = (
            f"这个 Excel 文件约 {size_mb:.1f} MB。\n\n"
            f"写回原 Excel 时可能临时占用约 {estimated_mb:.0f} MB 内存，并且需要几分钟。\n"
            "请先关闭 Excel、微信/网盘同步等可能占用文件的程序。\n\n"
            "是否继续？"
        )
        return messagebox.askyesno("大文件提示", message)

    def refresh_class_names(self) -> None:
        try:
            count = int(self.class_count_var.get().strip())
        except ValueError:
            return
        if count > 0:
            self.class_names_var.set("，".join(f"{idx}班" for idx in range(1, count + 1)))

    def choose_input(self) -> None:
        path = filedialog.askopenfilename(
            title="选择学生名单",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")],
        )
        if not path:
            return
        self.input_var.set(path)
        input_path = Path(path)
        if "分班结果" in input_path.stem:
            self.write_log("提醒：当前选择的文件名包含“分班结果”，请确认它不是上一次生成的结果表。")

    def create_template_file(self) -> None:
        path = filedialog.asksaveasfilename(
            title="保存名单模板",
            initialfile="分班名单模板.xlsx",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")],
        )
        if not path:
            return
        try:
            created = create_template(path)
        except Exception as exc:  # pragma: no cover - GUI safety net.
            messagebox.showerror("生成失败", str(exc))
            return
        self.write_log(f"已生成模板：{created}")
        messagebox.showinfo("完成", f"已生成模板：\n{created}")

    def parse_int(self, value: str, label: str) -> int | None:
        value = value.strip()
        if not value:
            return None
        try:
            return int(value)
        except ValueError as exc:
            raise ValueError(f"{label}必须是数字。") from exc

    def parse_and_validate_class_names(self, class_count: int | None) -> list[str] | None:
        if class_count is not None and class_count < 1:
            messagebox.showwarning("参数错误", "班级数量必须大于0。")
            return None

        class_names = split_csv_like(self.class_names_var.get())
        if not class_names:
            return []

        self.class_names_var.set("，".join(class_names))

        if len(set(class_names)) != len(class_names):
            messagebox.showwarning("参数错误", "班级名称不能重复，请检查后再开始分班。")
            return None

        if class_count is None:
            return class_names

        if len(class_names) == class_count:
            return class_names

        if len(class_names) < class_count:
            missing = class_count - len(class_names)
            next_names = "，".join(f"{idx}班" for idx in range(len(class_names) + 1, class_count + 1))
            message = (
                f"班级数量填了 {class_count}，但班级名称只有 {len(class_names)} 个。\n\n"
                f"还需要增加 {missing} 个班级名称。"
            )
            if next_names:
                message += f"\n\n可以在最后补：，{next_names}"
            message += "\n\n班级名称请用中文逗号“，”隔开。"
        else:
            extra = len(class_names) - class_count
            message = (
                f"班级数量填了 {class_count}，但班级名称有 {len(class_names)} 个。\n\n"
                f"请删去多出的 {extra} 个班级名称，或把班级数量改为 {len(class_names)}。\n\n"
                "班级名称请用中文逗号“，”隔开。"
            )
        messagebox.showwarning("班级数量不匹配", message)
        return None

    def start_split(self) -> None:
        if self.is_running:
            self.write_log("分班正在进行中，请等待当前任务完成。")
            return
        input_path = self.input_var.get().strip()
        if not input_path:
            messagebox.showwarning("缺少文件", "请先选择名单文件。")
            return

        try:
            class_count = self.parse_int(self.class_count_var.get(), "班级数量")
            seed = self.parse_int(self.seed_var.get(), "随机种子")
            capacity = self.parse_int(self.capacity_var.get(), "每班上限")
        except ValueError as exc:
            messagebox.showwarning("参数错误", str(exc))
            return

        class_names = self.parse_and_validate_class_names(class_count)
        if class_names is None:
            return

        if not self.warn_large_file_if_needed(input_path):
            self.write_log("已取消分班。")
            return

        balance_options = self.balance_var.get().strip().upper() or "ABC"
        invalid = [char for char in balance_options if char.isalpha() and char not in BALANCE_OPTIONS]
        if invalid:
            messagebox.showwarning("参数错误", "分班依据只能输入 A、B、C、D，例如 A、AB、ABC、ABCD。")
            return
        self.set_running(True)
        self.write_log("开始分班...")
        threading.Thread(
            target=self._run_split,
            args=(input_path, class_count, class_names, seed, capacity, balance_options),
            daemon=True,
        ).start()

    def _run_split(
        self,
        input_path: str,
        class_count: int | None,
        class_names: list[str],
        seed: int | None,
        capacity: int | None,
        balance_options: str,
    ) -> None:
        try:
            result = split_workbook(
                input_path=input_path,
                class_count=class_count,
                class_names=class_names or None,
                seed=seed,
                capacity=capacity,
                balance_options=balance_options,
            )
        except ClassSplitError as exc:
            self.after(0, lambda: self.write_log(str(exc)))
            self.after(0, lambda: messagebox.showerror("分班失败", str(exc)))
            self.after(0, lambda: self.set_running(False))
            return
        except Exception as exc:  # pragma: no cover - GUI safety net.
            self.after(0, lambda: self.write_log(f"发生错误：{exc}"))
            self.after(0, lambda: messagebox.showerror("分班失败", str(exc)))
            self.after(0, lambda: self.set_running(False))
            return

        def done() -> None:
            self.write_log(f"分班完成：已写入 {result.output_path} 的“分班结果”页面")
            self.write_log(f"学生人数：{result.student_count}；随机种子：{result.seed}；分班依据：{result.balance_options}")
            self.set_running(False)
            messagebox.showinfo("完成", f"分班完成，已在原 Excel 中新增/更新“分班结果”页面：\n{result.output_path}")

        self.after(0, done)


if __name__ == "__main__":
    SplitterApp().mainloop()
