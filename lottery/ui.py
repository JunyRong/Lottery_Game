"""抽奖软件图形界面（Tkinter / ttk）。"""

from __future__ import annotations

import os
import random
import tkinter as tk
import tkinter.font as tkfont
from collections import Counter
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from .config import load_tiers, save_tiers
from .engine import LotteryEngine, LotteryError
from .excel_io import (
    SheetPreview,
    default_column,
    export_results,
    inspect_sheet,
    list_sheets,
)
from .models import Entrant, PrizeTier, Winner
from .names import parse_manual_names
from .tier_io import TierPreview, read_tiers

BG = "#eaf3fc"
CARD = "#ffffff"
BORDER = "#cfe1f5"
INK = "#1e293b"
MUTED = "#64748b"
ACCENT = "#dc2626"
ACCENT_DARK = "#b91c1c"
STAGE_BG = "#0a2540"
STAGE_BORDER = "#1d4ed8"
STAGE_IDLE = "#7dd3fc"
STAGE_HINT = "#9cc3e5"
RESULT_RED = "#ff4d4f"
RESULT_RED_SOFT = "#ffb3b3"

FONT = ("Microsoft YaHei UI", 10)
FONT_BOLD = ("Microsoft YaHei UI", 10, "bold")
FONT_TITLE = ("Microsoft YaHei UI", 13, "bold")
FONT_STAGE = ("Microsoft YaHei UI", 58, "bold")
FONT_BUTTON = ("Microsoft YaHei UI", 15, "bold")

# 舞台字号从大到小尝试；批量结果会自动排成多列以容纳全部中奖者。
STAGE_FONT_SIZES = (
    58,
    52,
    46,
    40,
    34,
    28,
    24,
    20,
    18,
    16,
    14,
    12,
    10,
    9,
    8,
    7,
    6,
    5,
    4,
)


class SheetDialog(tk.Toplevel):
    """“选择工作表”弹窗的公共骨架：文件卡片、工作表下拉框与底部按钮。

    子类通过 `_build_content()` 追加自己的控件，并实现 `_load_sheet()` /
    `_on_ok()`；`result` 的具体类型由子类定义。
    """

    title_text = ""

    def __init__(self, master: tk.Misc, path: str):
        super().__init__(master)
        self.title(self.title_text)
        self.resizable(False, False)
        self.transient(master)
        self.path = path
        self.result = None

        try:
            sheets = list_sheets(path)
        except Exception as exc:  # openpyxl 的异常类型较杂，统一转为可读提示
            self.destroy()
            raise RuntimeError(str(exc)) from exc

        card = tk.Frame(self, bg=CARD, padx=20, pady=16)
        card.pack(fill=tk.BOTH, expand=True)

        tk.Label(
            card,
            text=f"文件：{Path(path).name}",
            bg=CARD,
            fg=INK,
            font=FONT_BOLD,
            anchor="w",
            wraplength=560,
        ).pack(fill=tk.X, pady=(0, 12))

        form = tk.Frame(card, bg=CARD)
        form.pack(fill=tk.X)
        tk.Label(form, text="工作表", bg=CARD, fg=INK, font=FONT).grid(
            row=0, column=0, sticky="w", pady=4
        )
        self.sheet_combo = ttk.Combobox(
            form, values=sheets, state="readonly", width=40, font=FONT
        )
        self.sheet_combo.grid(row=0, column=1, sticky="ew", padx=(10, 0), pady=4)
        self.sheet_combo.current(0)
        self.sheet_combo.bind("<<ComboboxSelected>>", lambda _e: self._load_sheet())
        form.columnconfigure(1, weight=1)

        self._build_content(card, form)

        buttons = tk.Frame(card, bg=CARD)
        buttons.pack(fill=tk.X, pady=(12, 0))
        self.ok_button = ttk.Button(buttons, text="导入", command=self._on_ok)
        self.ok_button.pack(side=tk.RIGHT)
        ttk.Button(buttons, text="取消", command=self.destroy).pack(
            side=tk.RIGHT, padx=(0, 8)
        )

        self.bind("<Return>", lambda _e: self._on_ok())
        self.bind("<Escape>", lambda _e: self.destroy())
        self._load_sheet()
        self.grab_set()
        self.focus_set()

    def _build_content(self, card: tk.Frame, form: tk.Frame) -> None:
        raise NotImplementedError

    def _load_sheet(self) -> None:
        raise NotImplementedError

    def _on_ok(self) -> None:
        raise NotImplementedError


class ImportDialog(SheetDialog):
    """选择工作表与姓名列。result = (names, sheet_name) 或 None。"""

    title_text = "导入抽奖名单"

    def __init__(self, master: tk.Misc, path: str):
        self.preview: SheetPreview | None = None
        self.result: tuple[list[str], str] | None = None
        super().__init__(master, path)

    def _build_content(self, card: tk.Frame, form: tk.Frame) -> None:
        tk.Label(form, text="姓名列", bg=CARD, fg=INK, font=FONT).grid(
            row=1, column=0, sticky="w", pady=4
        )
        self.column_combo = ttk.Combobox(
            form, values=[], state="readonly", width=40, font=FONT
        )
        self.column_combo.grid(row=1, column=1, sticky="ew", padx=(10, 0), pady=4)
        self.column_combo.bind("<<ComboboxSelected>>", lambda _e: self._update_count())

        self.info_var = tk.StringVar(value="正在读取…")
        tk.Label(
            card,
            textvariable=self.info_var,
            bg=CARD,
            fg=MUTED,
            font=FONT,
            anchor="w",
            justify="left",
            wraplength=560,
        ).pack(fill=tk.X, pady=(12, 4))

    def _load_sheet(self) -> None:
        sheet = self.sheet_combo.get()
        try:
            preview = inspect_sheet(self.path, sheet)
        except Exception as exc:
            self.preview = None
            self.column_combo["values"] = []
            self.column_combo.set("")
            self.info_var.set("读取失败，请选择其他工作表。")
            self.ok_button.state(["disabled"])
            messagebox.showerror("读取失败", str(exc), parent=self)
            return
        self.preview = preview
        if not preview.columns:
            self.column_combo["values"] = []
            self.column_combo.set("")
            self.info_var.set("该工作表没有可用数据。")
            self.ok_button.state(["disabled"])
            return
        self.column_combo["values"] = [col.label for col in preview.columns]
        default = default_column(preview)
        if default is not None:
            self.column_combo.current(default)
        self.ok_button.state(["!disabled"])
        self._update_count()

    def _update_count(self) -> None:
        if self.preview is None:
            return
        idx = self.column_combo.current()
        if not 0 <= idx < len(self.preview.columns):
            self.info_var.set("请选择姓名列。")
            self.ok_button.state(["disabled"])
            return
        count = len(self.preview.columns[idx].values)
        source = "已自动识别标题行" if self.preview.has_header else "未发现标题行，按数据表读取"
        self.info_var.set(
            f"将导入 {count} 条名单（{source}）。\n"
            "同名会保留为不同参与者，导入后按顺序分配编号。"
        )
        if count:
            self.ok_button.state(["!disabled"])
        else:
            self.ok_button.state(["disabled"])

    def _on_ok(self) -> None:
        if self.preview is None:
            return
        idx = self.column_combo.current()
        if not 0 <= idx < len(self.preview.columns):
            messagebox.showwarning("提示", "请选择包含姓名的列。", parent=self)
            return
        names = list(self.preview.columns[idx].values)
        if not names:
            messagebox.showwarning("提示", "该列没有可用姓名。", parent=self)
            return
        self.result = (names, self.preview.sheet)
        self.destroy()


class PrizeImportDialog(SheetDialog):
    """从 Excel 导入奖项设置。result 为 PrizeTier 列表或 None。"""

    title_text = "导入奖项设置"

    def __init__(self, master: tk.Misc, path: str):
        self.preview: TierPreview | None = None
        self.result: list[PrizeTier] | None = None
        super().__init__(master, path)

    def _build_content(self, card: tk.Frame, form: tk.Frame) -> None:
        self.info_var = tk.StringVar(value="正在读取…")
        tk.Label(
            card,
            textvariable=self.info_var,
            bg=CARD,
            fg=MUTED,
            font=FONT,
            anchor="w",
            justify="left",
            wraplength=560,
        ).pack(fill=tk.X, pady=(10, 6))

        tree_frame = tk.Frame(card, bg=CARD)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        self.tree = ttk.Treeview(
            tree_frame,
            columns=("name", "prize", "count"),
            show="headings",
            height=7,
        )
        for key, text, width, anchor in (
            ("name", "奖项", 150, "w"),
            ("prize", "奖品", 240, "w"),
            ("count", "人数", 70, "center"),
        ):
            self.tree.heading(key, text=text)
            self.tree.column(key, width=width, anchor=anchor)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=scroll.set)

        tk.Label(
            card,
            text="表头需包含“奖项/等级”、“奖品”、“人数/名额”。\n"
            "导入后按 Excel 行顺序排列；抽奖默认从列表末尾（最低奖项）开始。",
            bg=CARD,
            fg=MUTED,
            font=FONT,
            anchor="w",
            justify="left",
            wraplength=560,
        ).pack(fill=tk.X, pady=(8, 0))

    def _load_sheet(self) -> None:
        sheet = self.sheet_combo.get()
        try:
            preview = read_tiers(self.path, sheet)
        except Exception as exc:
            self.preview = None
            self.tree.delete(*self.tree.get_children())
            self.info_var.set(f"读取失败：{exc}")
            self.ok_button.state(["disabled"])
            return
        self.preview = preview
        self.tree.delete(*self.tree.get_children())
        for tier in preview.tiers:
            self.tree.insert("", "end", values=(tier.name, tier.prize, tier.count))
        source = "已自动识别标题行" if preview.has_header else "未发现标题行，按列顺序读取"
        self.info_var.set(f"识别到 {len(preview.tiers)} 个奖项（{source}）。")
        self.ok_button.state(["!disabled"])

    def _on_ok(self) -> None:
        if self.preview is None or not self.preview.tiers:
            return
        self.result = list(self.preview.tiers)
        self.destroy()


class ManualInputDialog(tk.Toplevel):
    """手动输入/编辑抽奖名单：逐人添加、批量粘贴、删除。

    弹窗会载入当前名单；result 为编辑后的完整姓名列表或 None。
    """

    def __init__(self, master: tk.Misc, initial_names: list[str] | None = None):
        super().__init__(master)
        self.title("手动输入 / 编辑名单")
        self.resizable(False, False)
        self.transient(master)
        self.result: list[str] | None = None
        self.names: list[str] = list(initial_names or [])

        card = tk.Frame(self, bg=CARD, padx=20, pady=16)
        card.pack(fill=tk.BOTH, expand=True)
        tk.Label(
            card,
            text="逐个添加抽奖人，或把姓名批量粘贴到下方。编号按列表顺序自动生成。",
            bg=CARD,
            fg=MUTED,
            font=FONT,
            anchor="w",
        ).pack(fill=tk.X, pady=(0, 8))

        # 单人添加
        add_row = tk.Frame(card, bg=CARD)
        add_row.pack(fill=tk.X, pady=(0, 10))
        self.name_entry = ttk.Entry(add_row, width=32, font=FONT)
        self.name_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.name_entry.bind("<Return>", lambda _e: self._add_single())
        ttk.Button(add_row, text="添加抽奖人", command=self._add_single).pack(
            side=tk.LEFT, padx=(8, 0)
        )

        # 当前名单
        tree_frame = tk.Frame(card, bg=CARD)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        self.tree = ttk.Treeview(
            tree_frame,
            columns=("number", "name"),
            show="headings",
            height=8,
            selectmode="extended",
        )
        self.tree.heading("number", text="编号")
        self.tree.heading("name", text="姓名")
        self.tree.column("number", width=70, anchor="center")
        self.tree.column("name", width=340, anchor="w")
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=scroll.set)

        list_actions = tk.Frame(card, bg=CARD)
        list_actions.pack(fill=tk.X, pady=(6, 10))
        ttk.Button(
            list_actions, text="删除选中", command=self._remove_selected
        ).pack(side=tk.LEFT)
        ttk.Button(list_actions, text="清空名单", command=self._clear_names).pack(
            side=tk.LEFT, padx=(8, 0)
        )

        # 批量添加
        tk.Label(
            card,
            text="批量添加（每行一个，也支持逗号、分号、顿号分隔）",
            bg=CARD,
            fg=INK,
            font=FONT,
            anchor="w",
        ).pack(fill=tk.X, pady=(0, 4))
        bulk_frame = tk.Frame(card, bg=CARD)
        bulk_frame.pack(fill=tk.X)
        self.bulk_text = tk.Text(
            bulk_frame,
            height=4,
            font=FONT,
            wrap="word",
            undo=True,
            padx=8,
            pady=6,
            relief="solid",
            borderwidth=1,
            highlightthickness=0,
        )
        self.bulk_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        bulk_scroll = ttk.Scrollbar(
            bulk_frame, orient=tk.VERTICAL, command=self.bulk_text.yview
        )
        bulk_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.bulk_text.configure(yscrollcommand=bulk_scroll.set)
        ttk.Button(card, text="批量添加", command=self._add_bulk).pack(
            anchor="w", pady=(6, 0)
        )

        footer = tk.Frame(card, bg=CARD)
        footer.pack(fill=tk.X, pady=(12, 0))
        self.count_var = tk.StringVar()
        tk.Label(
            footer, textvariable=self.count_var, bg=CARD, fg=MUTED, font=FONT
        ).pack(side=tk.LEFT)
        self.ok_button = ttk.Button(footer, text="确定", command=self._on_ok)
        self.ok_button.pack(side=tk.RIGHT)
        ttk.Button(footer, text="取消", command=self.destroy).pack(
            side=tk.RIGHT, padx=(0, 8)
        )

        self.bind("<Escape>", lambda _e: self.destroy())
        self.bind("<Control-Return>", lambda _e: self._on_ok())
        self._refresh_names()
        self.grab_set()
        self.name_entry.focus_set()

    def _refresh_names(self) -> None:
        self.tree.delete(*self.tree.get_children())
        for index, name in enumerate(self.names, start=1):
            self.tree.insert("", "end", iid=str(index - 1), values=(index, name))
        count = len(self.names)
        if count:
            self.count_var.set(f"已添加 {count} 位抽奖人（自动编号 1–{count}）")
        else:
            self.count_var.set("尚未添加抽奖人")
        self.ok_button.state(["!disabled"] if count else ["disabled"])

    def _add_single(self) -> None:
        name = self.name_entry.get().strip()
        if not name:
            self.name_entry.focus_set()
            return
        self.names.append(name)
        self.name_entry.delete(0, tk.END)
        self._refresh_names()
        self.tree.see(str(len(self.names) - 1))
        self.name_entry.focus_set()

    def _add_bulk(self) -> None:
        names = parse_manual_names(self.bulk_text.get("1.0", "end"))
        if not names:
            messagebox.showinfo("提示", "请先粘贴或输入姓名。", parent=self)
            self.bulk_text.focus_set()
            return
        self.names.extend(names)
        self.bulk_text.delete("1.0", tk.END)
        self._refresh_names()
        self.tree.see(str(len(self.names) - 1))

    def _remove_selected(self) -> None:
        selection = self.tree.selection()
        if not selection:
            messagebox.showinfo("提示", "请先选择要删除的抽奖人。", parent=self)
            return
        for index in sorted((int(iid) for iid in selection), reverse=True):
            del self.names[index]
        self._refresh_names()

    def _clear_names(self) -> None:
        if not self.names:
            return
        if messagebox.askyesno(
            "确认", f"确定清空全部 {len(self.names)} 位抽奖人吗？", parent=self
        ):
            self.names.clear()
            self._refresh_names()

    def _on_ok(self) -> None:
        # 把尚未点“添加”的输入也一并纳入，避免误以为已保存
        pending = self.name_entry.get().strip()
        if pending:
            self.names.append(pending)
            self.name_entry.delete(0, tk.END)
        self.names.extend(parse_manual_names(self.bulk_text.get("1.0", "end")))
        if not self.names:
            messagebox.showwarning("提示", "请至少添加一位抽奖人。", parent=self)
            self.name_entry.focus_set()
            return
        self.result = list(self.names)
        self.destroy()


class TierDialog(tk.Toplevel):
    """新增 / 编辑奖项。result 为 PrizeTier 或 None。

    新增时可选择插入位置（tiers 为当前奖项列表，按高→低排列）。
    """

    def __init__(
        self,
        master: tk.Misc,
        tier: PrizeTier | None = None,
        tiers: list[PrizeTier] | None = None,
    ):
        super().__init__(master)
        self.title("编辑奖项" if tier else "添加奖项")
        self.resizable(False, False)
        self.transient(master)
        self.result: PrizeTier | None = None
        self.position: int | None = None
        self.tier = tier
        self._position_map: list[int] = []

        card = tk.Frame(self, bg=CARD, padx=20, pady=16)
        card.pack(fill=tk.BOTH, expand=True)
        self.name_var = tk.StringVar(value=tier.name if tier else "")
        self.prize_var = tk.StringVar(value=tier.prize if tier else "")
        self.count_var = tk.StringVar(value=str(tier.count) if tier else "1")

        rows = [
            ("奖项名称", self.name_var),
            ("奖品", self.prize_var),
            ("获奖人数", self.count_var),
        ]
        for row, (label, var) in enumerate(rows):
            tk.Label(card, text=label, bg=CARD, fg=INK, font=FONT).grid(
                row=row, column=0, sticky="w", pady=6
            )
            ttk.Entry(card, textvariable=var, width=28, font=FONT).grid(
                row=row, column=1, sticky="ew", padx=(12, 0), pady=6
            )
        card.columnconfigure(1, weight=1)

        row_count = len(rows)
        if tier is None and tiers:
            # 默认追加到末尾（最低奖项），也可插到任意奖项之前
            options = [
                f"第 {index + 1} 位 · 排在「{item.name}」之前"
                for index, item in enumerate(tiers)
            ]
            options.append("列表末尾（最低奖项）")
            self._position_map = [*range(len(tiers)), len(tiers)]
            self.position_var = tk.StringVar(value=options[-1])
            tk.Label(card, text="插入位置", bg=CARD, fg=INK, font=FONT).grid(
                row=row_count, column=0, sticky="w", pady=6
            )
            self.position_combo = ttk.Combobox(
                card,
                textvariable=self.position_var,
                values=options,
                state="readonly",
                width=26,
                font=FONT,
            )
            self.position_combo.current(len(options) - 1)
            self.position_combo.grid(
                row=row_count, column=1, sticky="ew", padx=(12, 0), pady=6
            )
            row_count += 1

        buttons = tk.Frame(card, bg=CARD)
        buttons.grid(row=row_count, column=0, columnspan=2, sticky="e", pady=(14, 0))
        ttk.Button(buttons, text="确定", command=self._on_ok).pack(side=tk.RIGHT)
        ttk.Button(buttons, text="取消", command=self.destroy).pack(
            side=tk.RIGHT, padx=(0, 8)
        )

        self.bind("<Return>", lambda _e: self._on_ok())
        self.bind("<Escape>", lambda _e: self.destroy())
        self.grab_set()
        self.focus_set()

    def _on_ok(self) -> None:
        name = self.name_var.get().strip()
        prize = self.prize_var.get().strip()
        if not name:
            messagebox.showwarning("提示", "请填写奖项名称。", parent=self)
            return
        try:
            count = int(self.count_var.get().strip())
        except ValueError:
            messagebox.showwarning("提示", "获奖人数必须是整数。", parent=self)
            return
        if count < 1:
            messagebox.showwarning("提示", "获奖人数必须大于 0。", parent=self)
            return
        try:
            if self.tier:
                self.result = PrizeTier(
                    name=name,
                    prize=prize or name,
                    count=count,
                    id=self.tier.id,
                )
            else:
                self.result = PrizeTier(name=name, prize=prize or name, count=count)
        except ValueError as exc:
            messagebox.showwarning("提示", str(exc), parent=self)
            return
        if self._position_map:
            index = self.position_combo.current()
            if 0 <= index < len(self._position_map):
                self.position = self._position_map[index]
        self.destroy()


class LotteryApp:
    ROLL_TICKS = 28
    ROLL_INTERVAL_MS = 55

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("幸运抽奖")
        self.root.geometry("1220x840")
        self.root.minsize(960, 640)
        self.root.configure(bg=BG)

        self.tiers: list[PrizeTier] = load_tiers()
        self.names: list[str] = []
        self.engine: LotteryEngine | None = None
        # 约定列表按“高→低”排列，默认从末尾（最低奖项）开始抽
        self.selected_tier_id: str | None = self.tiers[-1].id if self.tiers else None
        self.draw_mode = tk.StringVar(value="one")
        self._after_id: str | None = None
        self._busy = False
        self._result_visible = False
        self._stage_sub_hidden = False
        self._stage_sub_forced_hidden = False
        self._stage_winners: list[Winner] | None = None
        self._stage_winner_labels: list[tk.Label] = []
        self._stage_grid_layout: tuple[int, int] | None = None
        self._action_buttons: list[ttk.Button] = []

        self._build_styles()
        self._build_ui()
        self._refresh_all()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ---------------------------------------------------------------- 样式/布局
    def _build_styles(self) -> None:
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("TButton", font=FONT, padding=(12, 6))
        style.configure(
            "Accent.TButton",
            font=FONT_BUTTON,
            padding=(36, 10),
            background=ACCENT,
            foreground="#ffffff",
        )
        style.map(
            "Accent.TButton",
            background=[("active", ACCENT_DARK), ("disabled", "#cbd5e1")],
            foreground=[("disabled", "#f8fafc")],
        )
        style.configure("TCombobox", padding=4)
        style.map(
            "TCombobox",
            fieldbackground=[("readonly", "#f4f9ff")],
            foreground=[("readonly", INK)],
            selectbackground=[("readonly", "#1d4ed8")],
            selectforeground=[("readonly", "#ffffff")],
        )
        self.root.option_add("*TCombobox*Listbox.background", "#f8fbff")
        self.root.option_add("*TCombobox*Listbox.foreground", INK)
        self.root.option_add("*TCombobox*Listbox.selectBackground", "#1d4ed8")
        self.root.option_add("*TCombobox*Listbox.selectForeground", "#ffffff")
        style.configure("TRadiobutton", font=FONT, background=CARD)
        style.configure(
            "Treeview",
            font=FONT,
            rowheight=28,
            background=CARD,
            fieldbackground=CARD,
            borderwidth=0,
        )
        style.configure(
            "Treeview.Heading",
            font=FONT_BOLD,
            background="#f1f5f9",
            foreground=INK,
            relief="flat",
        )
        style.map(
            "Treeview",
            background=[("selected", "#1d4ed8")],
            foreground=[("selected", "#ffffff")],
        )

    def _card(self, parent: tk.Misc, **pack_kwargs) -> tk.Frame:
        card = tk.Frame(
            parent,
            bg=CARD,
            highlightbackground=BORDER,
            highlightthickness=1,
        )
        card.pack(**pack_kwargs)
        return card

    def _build_ui(self) -> None:
        outer = tk.Frame(self.root, bg=BG)
        outer.pack(fill=tk.BOTH, expand=True, padx=16, pady=10)

        # 顶部工具栏
        toolbar = self._card(outer, fill=tk.X, pady=(0, 8))
        bar = tk.Frame(toolbar, bg=CARD, padx=12, pady=6)
        bar.pack(fill=tk.X)

        import_btn = ttk.Button(bar, text="导入名单", command=self._on_import)
        import_btn.pack(side=tk.LEFT)
        manual_btn = ttk.Button(bar, text="手动输入", command=self._on_manual_input)
        manual_btn.pack(side=tk.LEFT, padx=(8, 0))
        export_btn = ttk.Button(bar, text="导出结果", command=self._on_export)
        export_btn.pack(side=tk.LEFT, padx=(8, 0))
        reset_btn = ttk.Button(bar, text="重置抽奖", command=self._on_reset)
        reset_btn.pack(side=tk.LEFT, padx=(8, 0))
        self._action_buttons.extend([import_btn, manual_btn, export_btn, reset_btn])

        self.file_var = tk.StringVar(value="尚未导入名单")
        tk.Label(bar, textvariable=self.file_var, bg=CARD, fg=MUTED, font=FONT).pack(
            side=tk.RIGHT
        )

        # 底部中奖名单先占位：窗口变小时由中间主体先收缩，
        # 避免抽奖按钮/中奖名单被挤出可视区域。
        winners_card = self._card(
            outer, side=tk.BOTTOM, fill=tk.X, pady=(8, 0)
        )
        self._build_winners(winners_card)

        # 主体：左奖项 / 右舞台
        body = tk.Frame(outer, bg=BG)
        body.pack(fill=tk.BOTH, expand=True)

        left = tk.Frame(
            body,
            bg=CARD,
            highlightbackground=BORDER,
            highlightthickness=1,
            width=400,
        )
        left.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 12))
        left.pack_propagate(False)
        self._build_tier_panel(left)

        right = tk.Frame(
            body,
            bg=CARD,
            highlightbackground=BORDER,
            highlightthickness=1,
        )
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self._build_stage(right)

    def _build_tier_panel(self, parent: tk.Frame) -> None:
        panel = tk.Frame(parent, bg=CARD, padx=14, pady=10)
        panel.pack(fill=tk.BOTH, expand=True)
        panel.grid_columnconfigure(0, weight=1)
        panel.grid_rowconfigure(2, weight=1, minsize=60)

        header = tk.Frame(panel, bg=CARD)
        header.grid(row=0, column=0, sticky="ew")
        tk.Label(header, text="奖项设置", bg=CARD, fg=INK, font=FONT_TITLE).pack(
            side=tk.LEFT
        )
        import_tiers_btn = ttk.Button(
            header, text="导入奖项", command=self._on_import_tiers
        )
        import_tiers_btn.pack(side=tk.RIGHT)
        self._action_buttons.append(import_tiers_btn)

        tk.Label(
            panel,
            text="双击奖项可编辑 · 列表按高→低排列，默认从最低奖项开始抽。"
            "添加时可选择插入位置，也可用「上移 / 下移」调整顺序。",
            bg=CARD,
            fg=MUTED,
            font=FONT,
            anchor="w",
            justify="left",
            wraplength=350,
        ).grid(row=1, column=0, sticky="ew", pady=(4, 8))

        tree_frame = tk.Frame(panel, bg=CARD)
        tree_frame.grid(row=2, column=0, sticky="nsew")
        columns = ("name", "prize", "count", "progress")
        self.tier_tree = ttk.Treeview(
            tree_frame, columns=columns, show="headings", height=10
        )
        headings = {
            "name": ("奖项", 90),
            "prize": ("奖品", 140),
            "count": ("人数", 50),
            "progress": ("进度", 65),
        }
        for key, (text, width) in headings.items():
            self.tier_tree.heading(key, text=text)
            self.tier_tree.column(key, width=width, anchor="center")
        self.tier_tree.column("prize", anchor="w")
        self.tier_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll = ttk.Scrollbar(
            tree_frame, orient=tk.VERTICAL, command=self.tier_tree.yview
        )
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.tier_tree.configure(yscrollcommand=scroll.set)
        self.tier_tree.bind("<<TreeviewSelect>>", self._on_tier_select)
        self.tier_tree.bind("<Double-1>", lambda _e: self._on_edit_tier())

        buttons = tk.Frame(panel, bg=CARD)
        buttons.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        add_btn = ttk.Button(buttons, text="添加奖项", command=self._on_add_tier)
        add_btn.pack(side=tk.LEFT)
        edit_btn = ttk.Button(buttons, text="编辑", command=self._on_edit_tier)
        edit_btn.pack(side=tk.LEFT, padx=(8, 0))
        del_btn = ttk.Button(buttons, text="删除", command=self._on_delete_tier)
        del_btn.pack(side=tk.LEFT, padx=(8, 0))
        self.move_down_btn = ttk.Button(
            buttons, text="下移", command=lambda: self._move_tier(1)
        )
        self.move_down_btn.pack(side=tk.RIGHT)
        self.move_up_btn = ttk.Button(
            buttons, text="上移", command=lambda: self._move_tier(-1)
        )
        self.move_up_btn.pack(side=tk.RIGHT, padx=(0, 8))
        self._action_buttons.extend(
            [add_btn, edit_btn, del_btn, self.move_up_btn, self.move_down_btn]
        )

    def _build_stage(self, parent: tk.Frame) -> None:
        panel = tk.Frame(parent, bg=CARD, padx=18, pady=10)
        panel.pack(fill=tk.BOTH, expand=True)
        panel.grid_columnconfigure(0, weight=1)
        panel.grid_rowconfigure(1, weight=1, minsize=90)

        header = tk.Frame(panel, bg=CARD)
        header.grid(row=0, column=0, sticky="ew")
        tk.Label(header, text="当前奖项", bg=CARD, fg=INK, font=FONT_TITLE).pack(
            side=tk.LEFT
        )
        self.tier_var = tk.StringVar()
        self.tier_combo = ttk.Combobox(
            header, textvariable=self.tier_var, state="readonly", width=30, font=FONT
        )
        self.tier_combo.pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=(12, 0)
        )
        self.tier_combo.bind("<<ComboboxSelected>>", self._on_combo_select)

        # 舞台行权重为 1：窗口变小时先压缩舞台，控制区始终完整可见。
        stage = tk.Frame(
            panel,
            bg=STAGE_BG,
            padx=20,
            pady=10,
            highlightbackground=STAGE_BORDER,
            highlightthickness=1,
        )
        stage.grid(row=1, column=0, sticky="nsew", pady=(8, 8))
        stage.grid_columnconfigure(0, weight=1)
        stage.grid_rowconfigure(0, weight=1)
        self.stage_box = stage

        self.stage_name = tk.StringVar(value="准备抽奖")
        self.stage_sub = tk.StringVar(value="导入名单后，点击下方按钮开始")
        self.stage_font = tkfont.Font(
            family=FONT_STAGE[0], size=FONT_STAGE[1], weight=FONT_STAGE[2]
        )
        self.stage_sub_font = tkfont.Font(family="Microsoft YaHei UI", size=11)

        # 舞台内容区：单人时显示一个大标签，多人时改为按可用空间自动排布网格。
        self.stage_body = tk.Frame(stage, bg=STAGE_BG)
        self.stage_body.grid(row=0, column=0, sticky="nsew")
        self.stage_body.bind("<Configure>", self._on_stage_resize)

        self.stage_label = tk.Label(
            self.stage_body,
            textvariable=self.stage_name,
            bg=STAGE_BG,
            fg=STAGE_IDLE,
            font=self.stage_font,
            wraplength=640,
            justify="center",
            anchor="center",
            padx=0,
            pady=0,
            borderwidth=0,
            highlightthickness=0,
        )
        self.stage_label.place(relx=0.5, rely=0.5, anchor="center")

        self.stage_grid = tk.Frame(self.stage_body, bg=STAGE_BG)

        self.stage_sub_label = tk.Label(
            stage,
            textvariable=self.stage_sub,
            bg=STAGE_BG,
            fg=STAGE_HINT,
            font=self.stage_sub_font,
            wraplength=680,
            justify="center",
            anchor="n",
        )
        self.stage_sub_label.grid(row=1, column=0, sticky="ew")

        controls = tk.Frame(panel, bg=CARD)
        controls.grid(row=2, column=0, sticky="ew")

        mode_row = tk.Frame(controls, bg=CARD)
        mode_row.pack(pady=(0, 2))
        tk.Label(mode_row, text="抽取方式：", bg=CARD, fg=MUTED, font=FONT).pack(
            side=tk.LEFT
        )
        ttk.Radiobutton(
            mode_row,
            text="抽 1 人",
            value="one",
            variable=self.draw_mode,
            command=self._update_draw_state,
        ).pack(side=tk.LEFT)
        self.all_radio = ttk.Radiobutton(
            mode_row,
            text="抽满该奖项",
            value="all",
            variable=self.draw_mode,
            command=self._update_draw_state,
        )
        self.all_radio.pack(side=tk.LEFT, padx=(10, 0))

        self.draw_button = ttk.Button(
            controls,
            text="开始抽奖",
            style="Accent.TButton",
            command=self._on_draw,
        )
        self.draw_button.pack(pady=(4, 6))
        self.info_var = tk.StringVar(value="名单：未导入 · 未中奖 0 人")
        tk.Label(
            controls,
            textvariable=self.info_var,
            bg=CARD,
            fg=MUTED,
            font=FONT,
            wraplength=520,
            justify="center",
        ).pack()

    # ------------------------------------------------------------ 舞台自适应
    def _on_stage_resize(self, event: tk.Event) -> None:
        """舞台尺寸变化时重新排布中奖者，保证全部内容都在可视区域内。"""
        if event.widget is not self.stage_body:
            return
        self.stage_label.configure(wraplength=max(160, event.width - 40))
        self.stage_sub_label.configure(wraplength=max(160, event.width))
        self._refresh_stage_layout()

    @staticmethod
    def _wrapped_lines(text: str, font: tkfont.Font, wrap: int) -> int:
        """估算文本在给定宽度下会占几行（用于给副标题预留高度）。"""
        lines = 0
        for paragraph in text.split("\n"):
            current = ""
            for char in paragraph:
                if current and font.measure(current + char) > wrap:
                    lines += 1
                    current = char
                else:
                    current += char
            lines += 1
        return max(1, lines)

    @staticmethod
    def _grid_gaps(size: int) -> tuple[int, int]:
        """列间距/行间距随字号缩小，尽量在有限空间里多放几列。"""
        if size >= 14:
            return (18, 3)
        if size <= 6:
            return (6, 1)
        if size <= 9:
            return (8, 1)
        return (10, 1)

    def _stage_available_height(self) -> tuple[int, int]:
        """返回舞台内容区的可用宽高，并按需显示或隐藏副标题。"""
        width = self.stage_body.winfo_width()
        if width <= 1:
            width = max(160, self.stage_box.winfo_width() - 40)
        total_height = max(24, self.stage_box.winfo_height() - 20)

        # 中奖者太多时强制让出副标题区域，优先保证全部姓名可见。
        if self._stage_sub_forced_hidden:
            self.stage_sub_label.grid_remove()
            self._stage_sub_hidden = True
            return width, max(24, total_height - 4)

        # 舞台过矮时先隐藏副标题，优先保证中奖者完整可见。
        sub_font = self.stage_sub_font
        if self.stage_sub_label.winfo_ismapped():
            sub_height = self.stage_sub_label.winfo_reqheight()
        else:
            sub_lines = self._wrapped_lines(self.stage_sub.get(), sub_font, width)
            sub_height = sub_lines * sub_font.metrics("linespace") + 6

        if self._stage_sub_hidden:
            if total_height - sub_height >= 36:
                self.stage_sub_label.grid()
                self._stage_sub_hidden = False
        elif total_height - sub_height < 28:
            self.stage_sub_label.grid_remove()
            self._stage_sub_hidden = True

        if self._stage_sub_hidden:
            return width, max(24, total_height - 4)
        return width, max(24, total_height - sub_height - 6)

    def _fit_stage_font(self, text: str) -> None:
        """按舞台大小缩放单人姓名；多人结果由 `_layout_stage_winners()` 处理。"""
        width, height = self._stage_available_height()
        wrap = max(160, width - 40)
        self.stage_label.configure(wraplength=wrap)
        for size in STAGE_FONT_SIZES:
            self.stage_font.configure(size=size)
            lines = self._wrapped_lines(text, self.stage_font, wrap)
            if lines * self.stage_font.metrics("linespace") <= height:
                break

    def _best_grid_fit(
        self, texts: list[str], width: int, height: int
    ) -> tuple[int, int, bool]:
        """选出放下全部中奖者的字号与列数，并说明是否真正放得下。"""
        count = len(texts)
        best_size, best_cols = STAGE_FONT_SIZES[-1], 1
        for size in STAGE_FONT_SIZES:
            self.stage_font.configure(size=size)
            line_height = self.stage_font.metrics("linespace")
            gap_x, gap_y = self._grid_gaps(size)
            cell_width = max(self.stage_font.measure(text) for text in texts) + gap_x
            cols = max(1, (width + gap_x) // cell_width)
            cols = min(cols, count)
            rows = (count + cols - 1) // cols
            content_height = rows * line_height + max(0, rows - 1) * gap_y
            best_size, best_cols = size, cols
            if content_height <= height:
                return size, cols, True
        return best_size, best_cols, False

    def _rebuild_stage_labels(self, texts: list[str], cols: int, size: int) -> None:
        """按行优先顺序把中奖者铺进网格，每行自动决定显示几人。"""
        for label in self._stage_winner_labels:
            label.destroy()
        self._stage_winner_labels.clear()

        gap_x, gap_y = self._grid_gaps(size)
        rows = (len(texts) + cols - 1) // cols
        for index, text in enumerate(texts):
            row, col = divmod(index, cols)
            label = tk.Label(
                self.stage_grid,
                text=text,
                bg=STAGE_BG,
                fg=RESULT_RED,
                font=self.stage_font,
                anchor="center",
                padx=0,
                pady=0,
                borderwidth=0,
                highlightthickness=0,
            )
            label.grid(
                row=row,
                column=col,
                sticky="n",
                padx=(0, gap_x if col < cols - 1 else 0),
                pady=(0, gap_y if row < rows - 1 else 0),
            )
            self._stage_winner_labels.append(label)

    def _layout_stage_winners(self) -> None:
        winners = self._stage_winners
        if not winners:
            return
        width, height = self._stage_available_height()
        texts = [f"{winner.number} 号 · {winner.name}" for winner in winners]
        size, cols, fits = self._best_grid_fit(texts, width, height)

        # 字号偏小时让出副标题区域往往能换来更大的字号，优先保证可读性。
        if not self._stage_sub_forced_hidden and not self._stage_sub_hidden:
            hidden_height = max(24, self.stage_box.winfo_height() - 20 - 4)
            alt_size, alt_cols, alt_fits = self._best_grid_fit(
                texts, width, hidden_height
            )
            prefer_alt = (not fits) or (fits and size < 12 and alt_size > size)
            if prefer_alt:
                self._stage_sub_forced_hidden = True
                self.stage_sub_label.grid_remove()
                self._stage_sub_hidden = True
                size, cols = alt_size, alt_cols

        self.stage_font.configure(size=size)
        signature = (size, cols)
        if signature == self._stage_grid_layout and len(self._stage_winner_labels) == len(
            texts
        ):
            return
        self._stage_grid_layout = signature
        self._rebuild_stage_labels(texts, cols, size)

    def _refresh_stage_layout(self) -> None:
        if self._stage_winners is None:
            self._fit_stage_font(self.stage_name.get())
        else:
            self._layout_stage_winners()

    def _set_stage_name(self, text: str) -> None:
        self._stage_winners = None
        self._stage_grid_layout = None
        self._stage_sub_forced_hidden = False
        self.stage_name.set(text)
        self.stage_grid.place_forget()
        self.stage_label.place(relx=0.5, rely=0.5, anchor="center")
        self._fit_stage_font(text)

    def _set_stage_winners(self, winners: list[Winner]) -> None:
        """批量结果：完整写入中奖名单，并用自适应网格全部展示。"""
        self._stage_winners = list(winners)
        self.stage_name.set(self._batch_stage_text(winners))
        self._stage_grid_layout = None
        self._stage_sub_forced_hidden = False
        self.stage_label.place_forget()
        self.stage_grid.place(relx=0.5, rely=0.5, anchor="center")
        self._layout_stage_winners()

    def _set_stage_result(self, is_result: bool) -> None:
        """抽奖结果用红色，等待/提示状态用科技蓝浅色。"""
        color = RESULT_RED if is_result else STAGE_IDLE
        self.stage_label.configure(fg=color)
        for label in self._stage_winner_labels:
            label.configure(fg=color)
        self.stage_sub_label.configure(
            fg=RESULT_RED_SOFT if is_result else STAGE_HINT
        )

    @staticmethod
    def _batch_stage_text(winners: list[Winner]) -> str:
        """批量抽奖时列出全部中奖者，沿用与单人一致的「编号 号 · 姓名」格式。"""
        return "\n".join(f"{winner.number} 号 · {winner.name}" for winner in winners)

    def _build_winners(self, parent: tk.Frame) -> None:
        panel = tk.Frame(parent, bg=CARD, padx=14, pady=10)
        panel.pack(fill=tk.BOTH, expand=True)

        header = tk.Frame(panel, bg=CARD)
        header.pack(fill=tk.X, pady=(0, 6))
        tk.Label(header, text="中奖名单", bg=CARD, fg=INK, font=FONT_TITLE).pack(
            side=tk.LEFT
        )
        self.winners_count_var = tk.StringVar(value="共 0 位")
        tk.Label(
            header, textvariable=self.winners_count_var, bg=CARD, fg=MUTED, font=FONT
        ).pack(side=tk.RIGHT)

        tree_frame = tk.Frame(panel, bg=CARD)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        columns = ("number", "tier", "prize", "name", "time")
        self.winners_tree = ttk.Treeview(
            tree_frame, columns=columns, show="headings", height=5
        )
        headings = {
            "number": ("编号", 65),
            "tier": ("奖项", 110),
            "prize": ("奖品", 210),
            "name": ("姓名", 180),
            "time": ("中奖时间", 160),
        }
        for key, (text, width) in headings.items():
            self.winners_tree.heading(key, text=text)
            self.winners_tree.column(key, width=width, anchor="center")
        self.winners_tree.column("name", anchor="w")
        self.winners_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll = ttk.Scrollbar(
            tree_frame, orient=tk.VERTICAL, command=self.winners_tree.yview
        )
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.winners_tree.configure(yscrollcommand=scroll.set)
        self.winners_tree.tag_configure("odd", background="#f8fafc")

    # ------------------------------------------------------------------ 刷新
    def _drawn_count(self, tier_id: str) -> int:
        return len(self.engine.winners_for(tier_id)) if self.engine else 0

    def _selected_index(self) -> int:
        """当前选中奖项的下标；未选中或已删除时返回 -1。"""
        for index, tier in enumerate(self.tiers):
            if tier.id == self.selected_tier_id:
                return index
        return -1

    def _current_tier(self) -> PrizeTier | None:
        index = self._selected_index()
        return self.tiers[index] if index >= 0 else None

    def _refresh_all(self) -> None:
        self._result_visible = False
        self._refresh_tiers()
        self._refresh_info()
        self._update_draw_state()

    def _refresh_tiers(self) -> None:
        self.tier_tree.delete(*self.tier_tree.get_children())
        # 一次统计所有奖项的中奖数，避免每个奖项都扫描一遍中奖名单。
        counts = Counter(w.tier_id for w in self.engine.winners) if self.engine else {}
        drawn = {tier.id: counts.get(tier.id, 0) for tier in self.tiers}
        if self.selected_tier_id not in drawn:
            self.selected_tier_id = self.tiers[-1].id if self.tiers else None

        for tier in self.tiers:
            self.tier_tree.insert(
                "",
                "end",
                iid=tier.id,
                values=(
                    tier.name,
                    tier.prize,
                    tier.count,
                    f"{drawn[tier.id]}/{tier.count}",
                ),
            )

        if self.selected_tier_id:
            self.tier_tree.selection_set(self.selected_tier_id)
            self.tier_tree.see(self.selected_tier_id)

        self.tier_combo["values"] = [
            f"{t.name} · {t.prize}（已抽 {drawn[t.id]}/{t.count}）"
            for t in self.tiers
        ]
        index = self._selected_index()
        if index >= 0:
            self.tier_combo.current(index)
        else:
            self.tier_combo.set("")
        self._update_move_buttons()
        self._update_draw_state()

    def _refresh_info(self) -> None:
        if not self.names:
            self.info_var.set("名单：未导入 · 未中奖 0 人")
            return
        remaining = self.engine.remaining_count if self.engine else 0
        self.info_var.set(
            f"名单：{len(self.names)} 人 · 未中奖 {remaining} 人 · 已中奖 {len(self.names) - remaining} 人"
        )

    def _update_draw_state(self) -> None:
        tier = self._current_tier()
        slots = 0
        full = False
        if tier and self.engine:
            drawn = self._drawn_count(tier.id)
            full = drawn >= tier.count
            slots = max(0, min(tier.count - drawn, self.engine.remaining_count))
        enabled = bool(
            not self._busy
            and tier is not None
            and self.engine is not None
            and self.engine.remaining_count > 0
            and not full
        )
        self.draw_button.state(["!disabled"] if enabled else ["disabled"])
        self.all_radio.configure(
            text=f"抽满该奖项（{slots} 人）" if slots else "抽满该奖项"
        )
        self.draw_button.configure(
            text="抽出全部获奖者"
            if self.draw_mode.get() == "all"
            else "开始抽奖"
        )
        if not self._busy and not self._result_visible:
            if not self.tiers:
                hint = "请先在左侧添加奖项"
            elif not self.names:
                hint = "请先导入抽奖名单"
            elif tier and self.engine and full:
                hint = f"「{tier.name}」已抽满，请选择其他奖项"
            elif self.engine and self.engine.remaining_count == 0:
                hint = "所有人都已中奖"
            elif tier:
                hint = f"点击下方按钮抽取「{tier.name}」"
            else:
                hint = ""
            if hint:
                self.stage_sub.set(hint)
            self._set_stage_result(False)
            self._refresh_stage_layout()

    def _on_tier_select(self, _event: tk.Event) -> None:
        selection = self.tier_tree.selection()
        if not selection:
            return
        self.selected_tier_id = selection[0]
        index = self._selected_index()
        if index >= 0:
            self.tier_combo.current(index)
        self._update_move_buttons()
        self._update_draw_state()

    def _on_combo_select(self, _event: tk.Event) -> None:
        index = self.tier_combo.current()
        if 0 <= index < len(self.tiers):
            self.selected_tier_id = self.tiers[index].id
            self.tier_tree.selection_set(self.selected_tier_id)
        self._update_move_buttons()
        self._update_draw_state()

    def _update_move_buttons(self) -> None:
        """上移/下移只在可编辑且有相邻奖项时可用。"""
        index = self._selected_index()
        self.move_up_btn.state(
            ["disabled"] if self._busy or index <= 0 else ["!disabled"]
        )
        self.move_down_btn.state(
            ["disabled"]
            if self._busy or index < 0 or index >= len(self.tiers) - 1
            else ["!disabled"]
        )

    # ------------------------------------------------------------------ 抽奖
    def _on_draw(self) -> None:
        if self._busy:
            return
        tier = self._current_tier()
        if tier is None:
            messagebox.showinfo("提示", "请先添加并选择一个奖项。")
            return
        if self.engine is None or not self.names:
            messagebox.showinfo("提示", "请先导入抽奖名单。")
            return
        if self.engine.is_full(tier.id):
            messagebox.showinfo("提示", f"「{tier.name}」已经抽满。")
            return
        if self.engine.remaining_count == 0:
            messagebox.showinfo("提示", "所有人都已中奖。")
            return

        self._set_busy(True)
        self._result_visible = False
        self._set_stage_result(False)
        self.stage_sub.set("正在抽取…")
        self._roll(tier, self.ROLL_TICKS, self.draw_mode.get() == "all")

    def _roll(self, tier: PrizeTier, ticks: int, draw_all: bool) -> None:
        if not self.engine:
            self._set_busy(False)
            return
        pool = self.engine.remaining_entrants
        if not pool:
            self._set_busy(False)
            self._update_draw_state()
            return
        preview = random.choice(pool)
        self._set_stage_name(f"{preview.number} 号 · {preview.name}")
        if ticks > 0:
            self._after_id = self.root.after(
                self.ROLL_INTERVAL_MS, self._roll, tier, ticks - 1, draw_all
            )
        else:
            self._finish_draw(tier, draw_all)

    def _finish_draw(self, tier: PrizeTier, draw_all: bool = False) -> None:
        self._after_id = None
        if not self.engine:
            self._set_busy(False)
            return
        remaining_slots = tier.count - len(self.engine.winners_for(tier.id))
        try:
            if draw_all:
                winners = self.engine.draw_all(tier.id)
            else:
                winners = [self.engine.draw(tier.id)]
        except LotteryError as exc:
            messagebox.showwarning("提示", str(exc))
            self._set_busy(False)
            self._update_draw_state()
            return
        if not winners:
            self._set_busy(False)
            self._update_draw_state()
            return

        if len(winners) == 1:
            self._set_stage_name(f"{winners[0].number} 号 · {winners[0].name}")
        else:
            self._set_stage_winners(winners)

        total = len(self.engine.winners)
        for winner in reversed(winners):
            self.winners_tree.insert(
                "",
                0,
                iid=f"winner-{winner.entrant.id}",
                values=(
                    winner.number,
                    winner.tier_name,
                    winner.prize,
                    winner.name,
                    winner.drawn_at.strftime("%H:%M:%S"),
                ),
                tags=("odd",) if total % 2 else (),
            )
            total -= 1
        self.winners_count_var.set(f"共 {len(self.engine.winners)} 位")
        self._set_busy(False)
        advanced = self._advance_after_draw(tier)
        self._refresh_tiers()
        self._refresh_info()

        # 放在刷新之后，避免被“奖项已抽满”等状态提示覆盖
        if len(winners) == 1:
            message = f"恭喜获得 {tier.name} · {tier.prize}"
        else:
            message = f"恭喜获得 {tier.name} · {tier.prize}（本轮 {len(winners)} 人）"
            if len(winners) < remaining_slots:
                message += "（名单不足）"
        if advanced:
            message += f" · 下一项「{advanced.name}」"
        self.stage_sub.set(message)
        self._set_stage_result(True)
        self._refresh_stage_layout()
        self._result_visible = True

    def _advance_after_draw(self, tier: PrizeTier) -> PrizeTier | None:
        """当前奖项抽满后，自动切换到上一项（更高奖项）。"""
        if not self.engine or not self.engine.is_full(tier.id):
            return None
        try:
            index = self.tiers.index(tier)
        except ValueError:
            return None
        for candidate in reversed(self.tiers[:index]):
            if not self.engine.is_full(candidate.id):
                self.selected_tier_id = candidate.id
                return candidate
        return None

    def _set_busy(self, busy: bool) -> None:
        self._busy = busy
        for button in self._action_buttons:
            button.state(["disabled"] if busy else ["!disabled"])
        self.tier_combo.configure(state=tk.DISABLED if busy else "readonly")
        self._update_move_buttons()
        self._update_draw_state()

    # ------------------------------------------------------------------ 导入/导出
    def _on_import(self) -> None:
        path = filedialog.askopenfilename(
            parent=self.root,
            title="选择抽奖名单",
            filetypes=[("Excel 工作簿", "*.xlsx *.xlsm"), ("所有文件", "*.*")],
        )
        if not path:
            return
        try:
            dialog = ImportDialog(self.root, path)
        except Exception as exc:
            messagebox.showerror("导入失败", f"无法读取该文件：\n{exc}")
            return
        self.root.wait_window(dialog)
        if dialog.result is None:
            return

        if self.engine and self.engine.winners:
            if not messagebox.askyesno(
                "确认", "当前已有中奖记录，导入新名单会清空抽奖结果。\n是否继续？"
            ):
                return

        names, sheet = dialog.result
        self._apply_names(names, f"{Path(path).name} · {sheet} · {len(names)} 人")

    def _on_manual_input(self) -> None:
        if self.engine and self.engine.winners:
            if not messagebox.askyesno(
                "确认", "当前已有中奖记录，修改名单会清空抽奖结果。\n是否继续？"
            ):
                return

        dialog = ManualInputDialog(self.root, initial_names=self.names)
        self.root.wait_window(dialog)
        if dialog.result is None or dialog.result == self.names:
            return
        self._apply_names(dialog.result, f"手动输入 · {len(dialog.result)} 人")

    def _apply_names(self, names: list[str], source_label: str) -> None:
        """替换名单并重置抽奖状态（Excel 导入与手动输入共用）。"""
        self.names = names
        self._rebuild_engine()
        self.file_var.set(source_label)
        self._reset_stage()

    def _on_export(self) -> None:
        if not self.engine or not self.engine.winners:
            messagebox.showinfo("提示", "还没有中奖记录可导出。")
            return
        default = f"抽奖结果_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
        path = filedialog.asksaveasfilename(
            parent=self.root,
            title="导出中奖结果",
            defaultextension=".xlsx",
            initialfile=default,
            filetypes=[("Excel 工作簿", "*.xlsx")],
        )
        if not path:
            return
        try:
            export_results(path, self.engine.winners, self.tiers)
        except PermissionError:
            messagebox.showerror("导出失败", "目标文件被占用，请关闭后重试。")
            return
        except Exception as exc:
            messagebox.showerror("导出失败", str(exc))
            return
        messagebox.showinfo("导出成功", f"已保存到：\n{path}")

    def _on_reset(self) -> None:
        if not self.engine or not self.engine.winners:
            messagebox.showinfo("提示", "当前没有中奖记录。")
            return
        if not messagebox.askyesno("确认", "确定清空所有中奖记录并重新开始吗？"):
            return
        self.engine.reset()
        self._reset_stage()

    # ------------------------------------------------------------------ 奖项维护
    def _on_import_tiers(self) -> None:
        if not self._tiers_editable():
            return
        path = filedialog.askopenfilename(
            parent=self.root,
            title="选择奖项设置 Excel",
            filetypes=[("Excel 工作簿", "*.xlsx *.xlsm"), ("所有文件", "*.*")],
        )
        if not path:
            return
        try:
            dialog = PrizeImportDialog(self.root, path)
        except Exception as exc:
            messagebox.showerror("导入失败", f"无法读取该文件：\n{exc}")
            return
        self.root.wait_window(dialog)
        if dialog.result is None:
            return
        if self.tiers and not messagebox.askyesno(
            "确认",
            f"导入将替换当前 {len(self.tiers)} 个奖项设置。\n是否继续？",
        ):
            return
        self.tiers = dialog.result
        self.selected_tier_id = self.tiers[-1].id
        save_tiers(self.tiers)
        self._rebuild_engine()
        self._refresh_all()

    def _tiers_editable(self) -> bool:
        if self.engine and self.engine.winners:
            messagebox.showinfo(
                "提示", "已有中奖记录，请先点击「重置抽奖」再修改奖项。"
            )
            return False
        return True

    def _on_add_tier(self) -> None:
        if not self._tiers_editable():
            return
        dialog = TierDialog(self.root, tiers=self.tiers)
        self.root.wait_window(dialog)
        if dialog.result is None:
            return
        position = dialog.position
        if position is None:
            position = len(self.tiers)
        position = max(0, min(position, len(self.tiers)))
        self.tiers.insert(position, dialog.result)
        self.selected_tier_id = dialog.result.id
        save_tiers(self.tiers)
        self._rebuild_engine()
        self._refresh_all()

    def _move_tier(self, offset: int) -> None:
        """上移/下移选中奖项，保持该项为当前奖项。"""
        if not self._tiers_editable():
            return
        tier = self._current_tier()
        if tier is None:
            messagebox.showinfo("提示", "请先选择一个奖项。")
            return
        index = self.tiers.index(tier)
        target = index + offset
        if not 0 <= target < len(self.tiers):
            return
        self.tiers[index], self.tiers[target] = self.tiers[target], self.tiers[index]
        self.selected_tier_id = tier.id
        save_tiers(self.tiers)
        self._rebuild_engine()
        self._refresh_all()

    def _on_edit_tier(self) -> None:
        if not self._tiers_editable():
            return
        tier = self._current_tier()
        if tier is None:
            messagebox.showinfo("提示", "请先选择一个奖项。")
            return
        dialog = TierDialog(self.root, tier)
        self.root.wait_window(dialog)
        if dialog.result is None:
            return
        index = self.tiers.index(tier)
        self.tiers[index] = dialog.result
        self.selected_tier_id = dialog.result.id
        save_tiers(self.tiers)
        self._rebuild_engine()
        self._refresh_all()

    def _on_delete_tier(self) -> None:
        if not self._tiers_editable():
            return
        tier = self._current_tier()
        if tier is None:
            messagebox.showinfo("提示", "请先选择一个奖项。")
            return
        if not messagebox.askyesno("确认", f"确定删除奖项「{tier.name}」吗？"):
            return
        self.tiers = [t for t in self.tiers if t.id != tier.id]
        self.selected_tier_id = self.tiers[-1].id if self.tiers else None
        save_tiers(self.tiers)
        self._rebuild_engine()
        self._refresh_all()

    # ------------------------------------------------------------------ 其他
    def _rebuild_engine(self) -> None:
        if not self.names:
            self.engine = None
            return
        entrants = [
            Entrant(id=i, number=i + 1, name=name)
            for i, name in enumerate(self.names)
        ]
        self.engine = LotteryEngine(entrants, self.tiers)

    def _clear_winners_view(self) -> None:
        self.winners_tree.delete(*self.winners_tree.get_children())
        self.winners_count_var.set("共 0 位")

    def _reset_stage(self) -> None:
        """清空中奖列表与舞台结果，回到“准备抽奖”状态。"""
        self._clear_winners_view()
        self._set_stage_name("准备抽奖")
        if self.tiers:
            self.selected_tier_id = self.tiers[-1].id
        self._refresh_all()

    def _on_close(self) -> None:
        if self._after_id is not None:
            try:
                self.root.after_cancel(self._after_id)
            except tk.TclError:
                pass
        self.root.destroy()


def run() -> None:
    if os.name == "nt":
        try:
            import ctypes

            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
    root = tk.Tk()
    LotteryApp(root)
    root.mainloop()
