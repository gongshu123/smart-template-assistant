from __future__ import annotations

import json
import os
import subprocess
import tempfile
import threading
import time
import argparse
import ctypes
import sys
from dataclasses import asdict
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from PIL import Image, ImageTk

from analyzer import FIELD_COLORS, FIELD_LABELS, FIELD_ORDER, LayerInfo, SlotMapping, TemplateAnalysis, analyze_template
from excel_parser import WorkbookData, parse_workbook, validate_assets
from photoshop_runner import build_jobs, find_photoshop, launch_batch, write_batch_script


APP_NAME = "智能套版助手"
APP_VERSION = "0.1.5 界面重构版"
ZOOM_MIN = 0.02
ZOOM_MAX = 4.0

BG = "#F3F6FA"
CARD = "#FFFFFF"
BORDER = "#D7DEE8"
TEXT = "#172033"
MUTED = "#667085"
PRIMARY = "#2563EB"
PRIMARY_DARK = "#1D4ED8"


def resource_path(relative_path: str) -> Path:
    root = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return root / relative_path


def enable_high_dpi():
    if os.name != "nt":
        return
    try:
        ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4))
        return
    except Exception:
        pass
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
        return
    except Exception:
        pass
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


class MappingGrid(tk.Frame):
    def __init__(self, parent, on_select):
        super().__init__(parent, bg=CARD, highlightbackground=BORDER, highlightthickness=1)
        self.on_select = on_select
        self.selected: str | None = None
        self.row_cells: dict[str, list[tk.Label]] = {}
        self.row_confidences: dict[str, str] = {}
        self.columnconfigure(0, weight=3)
        self.columnconfigure(1, weight=5)
        self.columnconfigure(2, weight=1)
        for column, (text, anchor) in enumerate((("Excel变量", "w"), ("PSD当前内容", "w"), ("状态", "center"))):
            label = tk.Label(
                self,
                text=text,
                bg="#EAF1FB",
                fg=TEXT,
                font=("Microsoft YaHei UI", 9, "bold"),
                anchor=anchor,
                relief="flat",
                borderwidth=0,
                highlightbackground=BORDER,
                highlightthickness=1,
                padx=7,
                pady=6,
            )
            label.grid(row=0, column=column, sticky="nsew")

    @staticmethod
    def _clip(value: str, limit: int) -> str:
        value = str(value)
        return value if len(value) <= limit else value[:limit - 1] + "…"

    def set_rows(self, rows):
        for cells in self.row_cells.values():
            for cell in cells:
                cell.destroy()
        self.row_cells.clear()
        self.row_confidences.clear()
        self.selected = None
        for row_index, (field, field_text, current_text, confidence) in enumerate(rows, start=1):
            self.row_confidences[field] = confidence
            values = (self._clip(field_text, 18), self._clip(current_text, 30), self._clip(confidence, 8))
            cells = []
            for column, value in enumerate(values):
                cell_bg = CARD if row_index % 2 else "#F8FAFC"
                cell_fg = TEXT
                if column == 2:
                    cell_bg, cell_fg = self._status_palette(confidence)
                cell = tk.Label(
                    self,
                    text=value,
                    bg=cell_bg,
                    fg=cell_fg,
                    font=("Microsoft YaHei UI", 9),
                    anchor="center" if column == 2 else "w",
                    relief="flat",
                    borderwidth=0,
                    highlightbackground=BORDER,
                    highlightthickness=1,
                    padx=7,
                    pady=6,
                )
                cell.grid(row=row_index, column=column, sticky="nsew")
                cell.bind("<Button-1>", lambda _event, selected_field=field: self.select(selected_field, notify=True))
                cells.append(cell)
            self.row_cells[field] = cells

    @staticmethod
    def _status_palette(confidence: str) -> tuple[str, str]:
        if confidence in {"高", "已保存"}:
            return "#DCFCE7", "#166534"
        if confidence in {"中"}:
            return "#FEF3C7", "#92400E"
        if confidence in {"手动确认"}:
            return "#DBEAFE", "#1D4ED8"
        return "#FEE2E2", "#B91C1C"

    def select(self, field: str, notify: bool = False):
        if field not in self.row_cells:
            return
        self.selected = field
        for row_index, (row_field, cells) in enumerate(self.row_cells.items(), start=1):
            background = "#DBEAFE" if row_field == field else (CARD if row_index % 2 else "#F8FAFC")
            foreground = "#1E40AF" if row_field == field else TEXT
            for column, cell in enumerate(cells):
                if column == 2:
                    status_bg, status_fg = self._status_palette(self.row_confidences.get(row_field, ""))
                    cell.configure(bg=status_bg, fg=status_fg)
                else:
                    cell.configure(bg=background, fg=foreground)
        if notify:
            self.on_select(field)


class SmartTemplateApp(tk.Tk):
    def __init__(self):
        super().__init__()
        try:
            self.tk.call("tk", "scaling", self.winfo_fpixels("1i") / 72.0)
        except tk.TclError:
            pass
        self.title(f"{APP_NAME} · {APP_VERSION}")
        self.geometry("1880x1735")
        self.minsize(1280, 760)
        self.configure(bg=BG)
        self.window_icon_photo: ImageTk.PhotoImage | None = None
        self.brand_photo: ImageTk.PhotoImage | None = None
        self._load_app_icon()

        self.analysis: TemplateAnalysis | None = None
        self.workbook_data: WorkbookData | None = None
        self.preview_image: Image.Image | None = None
        self.preview_photo: ImageTk.PhotoImage | None = None
        self.preview_scale = 1.0
        self.preview_fit_width = True
        self._fit_after_id: str | None = None
        self.selected_field: str | None = None
        self.highlight_layer_id: int | None = None
        self.batch_process: subprocess.Popen | None = None
        self.batch_log_path: Path | None = None
        self.batch_log_position = 0

        self.template_var = tk.StringVar()
        self.excel_var = tk.StringVar()
        self.assets_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.slot_var = tk.StringVar()
        self.zoom_var = tk.StringVar(value="100%")
        self.show_candidates_var = tk.BooleanVar(value=False)
        self.status_var = tk.StringVar(value="请选择PSD/PSB模板，然后点击“分析模板”")
        self.template_info_var = tk.StringVar(value="尚未分析")
        self.excel_info_var = tk.StringVar(value="尚未读取Excel")
        self.file_summary_var = tk.StringVar(value="等待选择模板、Excel和素材文件夹")
        self.source_collapsed = False
        self.photoshop_path = find_photoshop()

        self._configure_styles()
        self._build_ui()
        for variable in (self.template_var, self.excel_var, self.assets_var, self.output_var):
            variable.trace_add("write", self._update_file_summary)
        self._update_file_summary()
        self.after(80, self._set_opening_size)

    def _set_opening_size(self):
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        width = min(1880, max(1280, screen_w - 80))
        height = min(1735, max(760, screen_h - 120))
        x = max(0, (screen_w - width) // 2)
        y = max(0, (screen_h - height) // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")

    def _load_app_icon(self):
        icon_path = resource_path("assets/app_icon.png")
        if not icon_path.exists():
            return
        icon = Image.open(icon_path).convert("RGBA")
        self.window_icon_photo = ImageTk.PhotoImage(icon.resize((64, 64), Image.Resampling.LANCZOS))
        self.iconphoto(True, self.window_icon_photo)
        self.brand_photo = ImageTk.PhotoImage(icon.resize((112, 112), Image.Resampling.LANCZOS))

    def _configure_styles(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure(".", font=("Microsoft YaHei UI", 9), foreground=TEXT)
        style.configure("TFrame", background=CARD)
        style.configure("TLabel", background=CARD, foreground=TEXT)
        style.configure("Title.TLabel", font=("Microsoft YaHei UI", 18, "bold"), foreground=TEXT, background=BG)
        style.configure("Sub.TLabel", font=("Microsoft YaHei UI", 9), foreground=MUTED, background=BG)
        style.configure("Section.TLabelframe", background=CARD, bordercolor=BORDER, borderwidth=1, relief="solid")
        style.configure("Section.TLabelframe.Label", font=("Microsoft YaHei UI", 10, "bold"), foreground=TEXT, background=CARD)
        style.configure("Primary.TButton", font=("Microsoft YaHei UI", 10, "bold"), padding=(16, 8), foreground="#FFFFFF", background=PRIMARY, bordercolor=PRIMARY)
        style.map("Primary.TButton", background=[("active", PRIMARY_DARK), ("pressed", "#1E40AF")], foreground=[("disabled", "#CBD5E1"), ("!disabled", "#FFFFFF")])
        style.configure("Normal.TButton", font=("Microsoft YaHei UI", 9), padding=(10, 6), foreground=TEXT, background="#FFFFFF", bordercolor="#C8D1DC")
        style.map("Normal.TButton", background=[("active", "#EEF4FF"), ("pressed", "#E2EAF5")], bordercolor=[("active", "#93B4E7")])
        style.configure("Compact.TButton", font=("Microsoft YaHei UI", 9), padding=(7, 4), foreground=TEXT, background="#FFFFFF", bordercolor="#C8D1DC")
        style.map("Compact.TButton", background=[("active", "#EEF4FF"), ("pressed", "#E2EAF5")])
        style.configure("Toolbar.TButton", font=("Microsoft YaHei UI", 10, "bold"), padding=(7, 4), foreground="#344054", background="#F8FAFC", bordercolor="#D0D7E2")
        style.map("Toolbar.TButton", background=[("active", "#EAF1FB"), ("pressed", "#DBEAFE")])
        style.configure("TEntry", padding=(7, 6), fieldbackground="#FFFFFF", bordercolor="#C8D1DC")
        style.configure("TCombobox", padding=(6, 5), fieldbackground="#FFFFFF", bordercolor="#C8D1DC")
        style.configure("TCheckbutton", background=CARD, foreground=TEXT)
        style.map("TCheckbutton", background=[("active", CARD)])
        style.configure("Treeview", font=("Microsoft YaHei UI", 9), rowheight=31)
        style.configure("Treeview.Heading", font=("Microsoft YaHei UI", 9, "bold"), background="#EAF1FB", foreground=TEXT)
        style.configure("TNotebook", background=CARD, borderwidth=0)
        style.configure("TNotebook.Tab", font=("Microsoft YaHei UI", 9), padding=(14, 7), background="#EEF2F7", foreground=MUTED)
        style.map("TNotebook.Tab", background=[("selected", "#FFFFFF")], foreground=[("selected", PRIMARY)])

    def _build_ui(self):
        header = tk.Frame(self, bg=BG)
        header.pack(fill="x", padx=20, pady=(14, 9))
        ttk.Label(header, text=APP_NAME, style="Title.TLabel").pack(side="left")
        subtitle = "本地运行 · 不上传文件 · Photoshop负责保存分层PSD"
        ttk.Label(header, text=subtitle, style="Sub.TLabel").pack(side="left", padx=(14, 0), pady=(8, 0))
        ps_text = f"Photoshop：{self.photoshop_path.parent.name}" if self.photoshop_path else "未检测到Photoshop"
        ps_bg = "#DCFCE7" if self.photoshop_path else "#FEE2E2"
        ps_fg = "#166534" if self.photoshop_path else "#B91C1C"
        tk.Label(header, text=("● " + ps_text), bg=ps_bg, fg=ps_fg, font=("Microsoft YaHei UI", 9, "bold"), padx=11, pady=5).pack(side="right", pady=(4, 0))

        self.source_frame = ttk.LabelFrame(self, text="① 准备文件", style="Section.TLabelframe")
        self.source_frame.pack(fill="x", padx=20, pady=(0, 11))

        self.source_summary = tk.Frame(self.source_frame, bg=CARD)
        tk.Label(
            self.source_summary,
            textvariable=self.file_summary_var,
            bg=CARD,
            fg="#344054",
            font=("Microsoft YaHei UI", 9),
            anchor="w",
        ).pack(side="left", fill="x", expand=True, padx=14, pady=9)
        ttk.Button(self.source_summary, text="修改文件", style="Compact.TButton", command=lambda: self.toggle_source(False)).pack(side="right", padx=10, pady=6)

        self.source_body = tk.Frame(self.source_frame, bg=CARD)
        self.source_body.pack(fill="x")
        self.source_body.columnconfigure(1, weight=0)
        self.source_body.columnconfigure(4, weight=1)
        self._file_row(self.source_body, 0, "PSD/PSB模板", self.template_var, self.choose_template)
        self._file_row(self.source_body, 1, "Excel数据", self.excel_var, self.choose_excel)
        self._file_row(self.source_body, 2, "产品素材文件夹", self.assets_var, self.choose_assets, folder=True)
        self._file_row(self.source_body, 3, "PSD输出文件夹", self.output_var, self.choose_output, folder=True)
        action_frame = tk.Frame(self.source_body, bg=CARD)
        action_frame.grid(row=0, column=3, rowspan=4, sticky="ns", padx=(12, 10), pady=8)
        ttk.Button(action_frame, text="分析模板", style="Primary.TButton", command=self.start_analysis).pack(fill="x", pady=(0, 6))
        ttk.Button(action_frame, text="读取Excel", style="Normal.TButton", command=self.load_excel).pack(fill="x")
        ttk.Button(action_frame, text="收起文件区", style="Compact.TButton", command=lambda: self.toggle_source(True)).pack(fill="x", pady=(8, 0))
        brand_frame = tk.Frame(self.source_body, width=340, height=150, bg="#FFF8F3", highlightbackground="#E7D7CC", highlightthickness=1)
        brand_frame.grid(row=0, column=5, rowspan=4, sticky="nse", padx=(14, 12), pady=8)
        brand_frame.grid_propagate(False)
        if self.brand_photo:
            tk.Label(brand_frame, image=self.brand_photo, bg="#FFF8F3").pack(side="left", padx=(20, 12), pady=16)
        tk.Label(
            brand_frame,
            text="智能套版助手\n作者：巩树",
            bg="#FFF8F3",
            fg="#49352E",
            justify="left",
            font=("Microsoft YaHei UI", 13, "bold"),
            padx=4,
        ).pack(side="left", pady=16)

        content = tk.PanedWindow(self, orient=tk.HORIZONTAL, sashwidth=6, bg=BORDER, bd=0)
        content.pack(fill="both", expand=True, padx=20, pady=(0, 11))

        preview_frame = ttk.LabelFrame(content, text="② 画面变量确认", style="Section.TLabelframe")
        map_frame = ttk.LabelFrame(content, text="③ 商品位与变量映射", style="Section.TLabelframe")
        content.add(preview_frame, minsize=650, width=850)
        content.add(map_frame, minsize=360, width=450)

        preview_toolbar = tk.Frame(preview_frame, bg=CARD)
        preview_toolbar.pack(fill="x", padx=10, pady=8)
        ttk.Checkbutton(preview_toolbar, text="显示候选框", variable=self.show_candidates_var, command=self.redraw_preview).pack(side="left")
        ttk.Label(preview_toolbar, text="选择变量后，点击画面可重新绑定", background=CARD, foreground=MUTED).pack(side="left", padx=(12, 0))
        ttk.Button(preview_toolbar, text="适应宽度", style="Compact.TButton", command=self.fit_preview_width).pack(side="right")
        ttk.Button(preview_toolbar, text="+", style="Toolbar.TButton", command=lambda: self.adjust_preview_zoom(1.12), width=3).pack(side="right", padx=(5, 0))
        ttk.Button(preview_toolbar, textvariable=self.zoom_var, style="Compact.TButton", command=self.reset_preview_100, width=7).pack(side="right", padx=(5, 0))
        ttk.Button(preview_toolbar, text="−", style="Toolbar.TButton", command=lambda: self.adjust_preview_zoom(1 / 1.12), width=3).pack(side="right")

        canvas_wrap = tk.Frame(preview_frame, bg=CARD)
        canvas_wrap.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self.canvas = tk.Canvas(canvas_wrap, bg="#FFFFFF", highlightthickness=0)
        xscroll = ttk.Scrollbar(canvas_wrap, orient="horizontal", command=self.canvas.xview)
        yscroll = ttk.Scrollbar(canvas_wrap, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(xscrollcommand=xscroll.set, yscrollcommand=yscroll.set)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        canvas_wrap.rowconfigure(0, weight=1)
        canvas_wrap.columnconfigure(0, weight=1)
        self.canvas.bind("<Button-1>", self.on_canvas_click)
        self.canvas.bind("<MouseWheel>", self.on_preview_mousewheel)
        self.canvas.bind("<Configure>", self._on_canvas_configure)

        map_header = tk.Frame(map_frame, bg=CARD)
        map_header.pack(fill="x", padx=10, pady=(9, 6))
        ttk.Label(map_header, text="商品位", background="#FFFFFF").pack(side="left")
        self.slot_combo = ttk.Combobox(map_header, textvariable=self.slot_var, state="readonly", width=16)
        self.slot_combo.pack(side="left", padx=(7, 12))
        self.slot_combo.bind("<<ComboboxSelected>>", lambda _event: self.refresh_mapping())
        ttk.Label(map_header, textvariable=self.template_info_var, background="#FFFFFF", foreground="#636A73").pack(side="left")

        self.mapping_grid = MappingGrid(map_frame, self.select_mapping_field)
        self.mapping_grid.pack(fill="x", padx=10, pady=(0, 8))

        map_actions = tk.Frame(map_frame, bg=CARD)
        map_actions.pack(fill="x", padx=10, pady=(0, 8))
        ttk.Button(map_actions, text="定位选中变量", style="Normal.TButton", command=self.locate_selected).pack(side="left")
        ttk.Button(map_actions, text="清除绑定", style="Normal.TButton", command=self.clear_selected_binding).pack(side="left", padx=7)
        ttk.Button(map_actions, text="保存模板映射", style="Normal.TButton", command=self.save_mapping).pack(side="right")

        excel_info = tk.Frame(map_frame, bg="#F7F9FC", highlightbackground=BORDER, highlightthickness=1)
        excel_info.pack(fill="x", padx=10, pady=(0, 8))
        ttk.Label(excel_info, textvariable=self.excel_info_var, background="#F7F8FA", foreground="#4B535C", wraplength=380).pack(anchor="w", padx=9, pady=7)

        self.info_notebook = ttk.Notebook(map_frame)
        self.info_notebook.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        data_tab = tk.Frame(self.info_notebook, bg=CARD)
        activity_tab = tk.Frame(self.info_notebook, bg=CARD)
        self.info_notebook.add(data_tab, text="Excel数据预览")
        self.info_notebook.add(activity_tab, text="运行状态")

        data_columns = ("position", "name", "selling", "price", "image")
        self.data_tree = ttk.Treeview(data_tab, columns=data_columns, show="headings", selectmode="browse")
        headings = {
            "position": ("页面 / 商品位", 120),
            "name": ("产品名称", 150),
            "selling": ("产品卖点", 180),
            "price": ("价格", 80),
            "image": ("图片文件", 140),
        }
        for column, (title, width) in headings.items():
            self.data_tree.heading(column, text=title)
            self.data_tree.column(column, width=width, minwidth=60, anchor="w", stretch=column in {"name", "selling", "image"})
        data_scroll = ttk.Scrollbar(data_tab, orient="vertical", command=self.data_tree.yview)
        self.data_tree.configure(yscrollcommand=data_scroll.set)
        self.data_tree.pack(side="left", fill="both", expand=True)
        data_scroll.pack(side="right", fill="y")

        self.activity_text = tk.Text(
            activity_tab,
            bg="#F8FAFC",
            fg="#344054",
            relief="flat",
            borderwidth=0,
            wrap="word",
            font=("Microsoft YaHei UI", 9),
            padx=12,
            pady=10,
            state="disabled",
        )
        self.activity_text.pack(fill="both", expand=True)
        self.append_activity("请选择模板并点击“分析模板”。完成分析后，文件区域会自动收起。")

        footer = tk.Frame(self, bg=CARD, highlightbackground=BORDER, highlightthickness=1)
        footer.pack(fill="x", padx=20, pady=(0, 14))
        ttk.Label(footer, textvariable=self.status_var, background=CARD, foreground="#374151").pack(side="left", fill="x", expand=True, padx=12, pady=10)
        ttk.Button(footer, text="确认并批量生成PSD", style="Primary.TButton", command=lambda: self.run_batch(False)).pack(side="right", padx=(6, 10), pady=7)
        ttk.Button(footer, text="生成第1张测试稿", style="Normal.TButton", command=lambda: self.run_batch(True)).pack(side="right", padx=6, pady=7)
        ttk.Button(footer, text="打开输出目录", style="Normal.TButton", command=self.open_output_folder).pack(side="right", padx=6, pady=7)

    def toggle_source(self, collapse: bool | None = None):
        collapse = (not self.source_collapsed) if collapse is None else collapse
        if collapse == self.source_collapsed:
            return
        self.source_collapsed = collapse
        if collapse:
            self.source_body.pack_forget()
            self.source_summary.pack(fill="x")
        else:
            self.source_summary.pack_forget()
            self.source_body.pack(fill="x")

    def _update_file_summary(self, *_args):
        items = (
            ("模板", self.template_var.get()),
            ("Excel", self.excel_var.get()),
            ("素材", self.assets_var.get()),
            ("输出", self.output_var.get()),
        )
        self.file_summary_var.set("   |   ".join(f"{'✓' if value.strip() else '○'} {label}" for label, value in items))

    def append_activity(self, message: str):
        if not hasattr(self, "activity_text"):
            return
        timestamp = time.strftime("%H:%M:%S")
        self.activity_text.configure(state="normal")
        self.activity_text.insert("end", f"[{timestamp}] {message}\n")
        self.activity_text.see("end")
        self.activity_text.configure(state="disabled")

    def refresh_data_preview(self):
        if not hasattr(self, "data_tree"):
            return
        self.data_tree.delete(*self.data_tree.get_children())
        if not self.workbook_data:
            return
        for page in self.workbook_data.pages:
            for slot_index, product in enumerate(page.products, start=1):
                self.data_tree.insert(
                    "",
                    "end",
                    values=(
                        f"{page.name} / {slot_index}",
                        product.product_name,
                        product.selling_point,
                        product.price,
                        product.image,
                    ),
                )

    def _file_row(self, parent, row, label, variable, command, folder=False):
        ttk.Label(parent, text=label, background=CARD).grid(row=row, column=0, sticky="w", padx=(14, 10), pady=4)
        entry = ttk.Entry(parent, textvariable=variable, width=58)
        entry.grid(row=row, column=1, sticky="w", pady=4)
        ttk.Button(parent, text="选择文件夹" if folder else "选择文件", style="Normal.TButton", command=command).grid(row=row, column=2, padx=8, pady=4)

    def choose_template(self):
        path = filedialog.askopenfilename(title="选择PSD/PSB模板", filetypes=[("Photoshop模板", "*.psd *.psb"), ("所有文件", "*.*")])
        if path:
            self.template_var.set(path)
            if not self.output_var.get():
                self.output_var.set(str(Path(path).parent / "PSD输出"))

    def choose_excel(self):
        path = filedialog.askopenfilename(title="选择Excel数据", filetypes=[("Excel工作簿", "*.xlsx *.xlsm"), ("所有文件", "*.*")])
        if path:
            self.excel_var.set(path)

    def choose_assets(self):
        path = filedialog.askdirectory(title="选择产品素材文件夹")
        if path:
            self.assets_var.set(path)

    def choose_output(self):
        path = filedialog.askdirectory(title="选择PSD输出文件夹")
        if path:
            self.output_var.set(path)

    @property
    def app_data_dir(self) -> Path:
        base = Path(os.environ.get("LOCALAPPDATA", tempfile.gettempdir())) / "SmartTemplateAssistant"
        base.mkdir(parents=True, exist_ok=True)
        return base

    def mapping_path(self, fingerprint: str) -> Path:
        folder = self.app_data_dir / "mappings"
        folder.mkdir(parents=True, exist_ok=True)
        return folder / f"{fingerprint}.json"

    def start_analysis(self):
        template = Path(self.template_var.get().strip())
        if not template.exists():
            messagebox.showwarning(APP_NAME, "请先选择有效的PSD或PSB模板。")
            return
        self.status_var.set("正在分析可见图层和商品位，请稍候……")
        self.template_info_var.set("分析中")
        self.append_activity(f"开始分析模板：{template.name}")
        thread = threading.Thread(target=self._analysis_worker, args=(template,), daemon=True)
        thread.start()

    def _analysis_worker(self, template: Path):
        try:
            preview = self.app_data_dir / "preview.png"
            analysis = analyze_template(template, preview)
            self.after(0, lambda: self._analysis_done(analysis, preview))
        except Exception as error:
            self.after(0, lambda: self._show_error(f"模板分析失败：{error}"))

    def _analysis_done(self, analysis: TemplateAnalysis, preview_path: Path):
        self.analysis = analysis
        self._load_saved_mapping()
        self.preview_image = Image.open(preview_path).convert("RGBA")
        self.slot_combo["values"] = [f"商品位 {slot.slot_index}" for slot in analysis.slots]
        if analysis.slots:
            self.slot_combo.current(0)
        self.template_info_var.set(f"识别到 {len(analysis.slots)} 个商品位 · 跳过 {analysis.hidden_layers} 个隐藏图层")
        self.status_var.set("模板分析完成。请逐项点击右侧变量，核对左侧彩色框。")
        self.append_activity(f"模板分析完成：识别到{len(analysis.slots)}个商品位，跳过{analysis.hidden_layers}个隐藏图层。")
        self.refresh_mapping()
        self.after(150, lambda: self.toggle_source(True))
        self.after(100, self.fit_preview_width)
        if self.excel_var.get():
            self.load_excel()

    def _load_saved_mapping(self):
        if not self.analysis:
            return
        path = self.mapping_path(self.analysis.fingerprint)
        if not path.exists():
            return
        try:
            saved = json.loads(path.read_text(encoding="utf-8"))
            saved_slots = saved.get("slots", [])
            valid_ids = {layer.layer_id for layer in self.analysis.layers}
            for slot, saved_slot in zip(self.analysis.slots, saved_slots):
                for field, layer_id in saved_slot.get("fields", {}).items():
                    if field in slot.fields and layer_id in valid_ids:
                        slot.fields[field] = layer_id
                        slot.confidence[field] = "已保存"
        except Exception:
            pass

    def save_mapping(self):
        if not self.analysis:
            return
        data = {
            "template": self.analysis.template_path,
            "fingerprint": self.analysis.fingerprint,
            "slots": [{"slotIndex": slot.slot_index, "fields": slot.fields} for slot in self.analysis.slots],
        }
        self.mapping_path(self.analysis.fingerprint).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        self.status_var.set("模板映射已保存；下次使用此模板会自动载入。")
        self.append_activity("当前模板映射已保存。")

    def current_slot(self) -> SlotMapping | None:
        if not self.analysis or not self.analysis.slots:
            return None
        index = self.slot_combo.current()
        if index < 0:
            index = 0
        return self.analysis.slots[index]

    def layer_by_id(self, layer_id: int | None) -> LayerInfo | None:
        if not self.analysis or layer_id is None:
            return None
        return next((layer for layer in self.analysis.layers if layer.layer_id == layer_id), None)

    def layer_display(self, layer: LayerInfo | None) -> str:
        if not layer:
            return "— 未绑定 —"
        if layer.kind == "type":
            return (layer.text or layer.name).replace("\n", " / ")
        if layer.kind == "smartobject":
            return layer.smart_filename or layer.name
        return layer.name

    def refresh_mapping(self):
        slot = self.current_slot()
        if not slot:
            self.mapping_grid.set_rows([])
            self.redraw_preview()
            return
        rows = []
        for field in FIELD_ORDER:
            layer = self.layer_by_id(slot.fields.get(field))
            confidence = slot.confidence.get(field, "未识别")
            rows.append((field, FIELD_LABELS[field], self.layer_display(layer), confidence))
        self.mapping_grid.set_rows(rows)
        self.selected_field = None
        self.highlight_layer_id = None
        self.redraw_preview()

    def select_mapping_field(self, field: str):
        if field not in FIELD_ORDER:
            return
        self.selected_field = field
        self.mapping_grid.select(field)
        slot = self.current_slot()
        self.highlight_layer_id = slot.fields.get(self.selected_field) if slot else None
        self.redraw_preview()

    def locate_selected(self):
        if not self.selected_field:
            messagebox.showinfo(APP_NAME, "请先选择一个变量。")
            return
        self.select_mapping_field(self.selected_field)
        layer = self.layer_by_id(self.highlight_layer_id)
        if layer and layer.bbox:
            x = ((layer.bbox[0] + layer.bbox[2]) / 2) * self.preview_scale
            y = ((layer.bbox[1] + layer.bbox[3]) / 2) * self.preview_scale
            self.canvas.xview_moveto(max(0, (x - self.canvas.winfo_width() / 2) / max(1, self.preview_image.width * self.preview_scale)))
            self.canvas.yview_moveto(max(0, (y - self.canvas.winfo_height() / 2) / max(1, self.preview_image.height * self.preview_scale)))

    def clear_selected_binding(self):
        if not self.selected_field:
            messagebox.showinfo(APP_NAME, "请先选择一个变量。")
            return
        slot = self.current_slot()
        if slot:
            slot.fields[self.selected_field] = None
            slot.confidence[self.selected_field] = "未绑定"
            self.refresh_mapping()

    def fit_preview(self):
        if not self.preview_image:
            return
        self.update_idletasks()
        available_w = max(100, self.canvas.winfo_width() - 18)
        available_h = max(100, self.canvas.winfo_height() - 18)
        self.preview_scale = min(available_w / self.preview_image.width, available_h / self.preview_image.height)
        self.preview_scale = max(ZOOM_MIN, min(1.0, self.preview_scale))
        self.redraw_preview()
        self.canvas.xview_moveto(0)
        self.canvas.yview_moveto(0)

    def fit_preview_width(self):
        if not self.preview_image:
            return
        self.update_idletasks()
        available_w = max(100, self.canvas.winfo_width())
        self.preview_scale = max(ZOOM_MIN, min(ZOOM_MAX, available_w / self.preview_image.width))
        self.preview_fit_width = True
        self.redraw_preview()
        self.canvas.xview_moveto(0)
        self.canvas.yview_moveto(0)

    def reset_preview_100(self):
        if not self.preview_image:
            self.zoom_var.set("100%")
            return
        self.preview_scale = 1.0
        self.preview_fit_width = False
        self.redraw_preview()
        self.canvas.xview_moveto(0)
        self.canvas.yview_moveto(0)

    def adjust_preview_zoom(self, factor: float):
        if not self.preview_image:
            return
        self.preview_fit_width = False
        self.preview_scale = max(ZOOM_MIN, min(ZOOM_MAX, self.preview_scale * factor))
        self.redraw_preview()

    def _on_canvas_configure(self, _event):
        if not self.preview_image or not self.preview_fit_width:
            return
        if self._fit_after_id:
            self.after_cancel(self._fit_after_id)
        self._fit_after_id = self.after(80, self._refit_preview_after_resize)

    def _refit_preview_after_resize(self):
        self._fit_after_id = None
        self.fit_preview_width()

    def on_preview_mousewheel(self, event):
        control_pressed = bool(event.state & 0x0004)
        alt_pressed = bool(event.state & 0x0008)
        if not self.preview_image or not (control_pressed and alt_pressed) or event.delta == 0:
            return
        self.preview_fit_width = False
        old_scale = self.preview_scale
        image_x = self.canvas.canvasx(event.x) / old_scale
        image_y = self.canvas.canvasy(event.y) / old_scale
        steps = event.delta / 120.0
        self.preview_scale = max(ZOOM_MIN, min(ZOOM_MAX, old_scale * (1.12 ** steps)))
        if abs(self.preview_scale - old_scale) < 0.0001:
            return "break"
        self.redraw_preview()
        scaled_w = max(1, self.preview_image.width * self.preview_scale)
        scaled_h = max(1, self.preview_image.height * self.preview_scale)
        left = image_x * self.preview_scale - event.x
        top = image_y * self.preview_scale - event.y
        self.canvas.xview_moveto(max(0.0, min(1.0, left / scaled_w)))
        self.canvas.yview_moveto(max(0.0, min(1.0, top / scaled_h)))
        return "break"

    def redraw_preview(self):
        self.canvas.delete("all")
        if not self.preview_image:
            self.canvas.create_text(280, 250, text="分析模板后在这里显示彩色变量框", fill="#C5C9CE", font=("Microsoft YaHei UI", 13))
            self.zoom_var.set("100%")
            return
        self.zoom_var.set(f"{self.preview_scale * 100:.0f}%")
        size = (max(1, int(self.preview_image.width * self.preview_scale)), max(1, int(self.preview_image.height * self.preview_scale)))
        resized = self.preview_image.resize(size, Image.Resampling.LANCZOS)
        self.preview_photo = ImageTk.PhotoImage(resized)
        self.canvas.create_image(0, 0, image=self.preview_photo, anchor="nw")
        self.canvas.configure(scrollregion=(0, 0, size[0], size[1]))

        if not self.analysis:
            return
        slot = self.current_slot()
        if not slot:
            return

        if self.show_candidates_var.get():
            candidate_ids = {layer.layer_id for layer in self.analysis.layers if layer.visible and layer.bbox and layer.kind in {"type", "smartobject", "pixel"} and slot.root_layer_id in layer.path_ids}
            for layer_id in candidate_ids:
                layer = self.layer_by_id(layer_id)
                self._draw_box(layer, "#A7ADB4", "", 1, dash=(3, 3))

        for field in FIELD_ORDER:
            layer = self.layer_by_id(slot.fields.get(field))
            if layer:
                width = 5 if layer.layer_id == self.highlight_layer_id else 3
                self._draw_box(layer, FIELD_COLORS[field], FIELD_LABELS[field].split("（")[0], width)

    def _draw_box(self, layer: LayerInfo | None, color: str, label: str, width: int, dash=None):
        if not layer or not layer.bbox:
            return
        x1, y1, x2, y2 = [value * self.preview_scale for value in layer.bbox]
        self.canvas.create_rectangle(x1, y1, x2, y2, outline=color, width=width, dash=dash)
        if label:
            text_id = self.canvas.create_text(x1 + 4, max(9, y1 - 9), text=label, anchor="w", fill="#FFFFFF", font=("Microsoft YaHei UI", 8, "bold"))
            bounds = self.canvas.bbox(text_id)
            if bounds:
                bg = self.canvas.create_rectangle(bounds[0] - 3, bounds[1] - 1, bounds[2] + 3, bounds[3] + 1, fill=color, outline=color)
                self.canvas.tag_raise(text_id, bg)

    def on_canvas_click(self, event):
        if not self.analysis or not self.preview_image:
            return
        x = self.canvas.canvasx(event.x) / self.preview_scale
        y = self.canvas.canvasy(event.y) / self.preview_scale
        slot = self.current_slot()
        if not slot:
            return
        candidates = []
        wanted_kinds = {"smartobject", "pixel"} if self.selected_field == "image" else {"type"}
        for layer in self.analysis.layers:
            if not layer.visible or not layer.bbox or layer.kind not in wanted_kinds:
                continue
            if slot.root_layer_id not in layer.path_ids:
                continue
            x1, y1, x2, y2 = layer.bbox
            if x1 <= x <= x2 and y1 <= y <= y2:
                candidates.append(layer)
        if not candidates:
            return
        layer = min(candidates, key=lambda candidate: candidate.area)
        if self.selected_field:
            field = self.selected_field
            slot.fields[field] = layer.layer_id
            slot.confidence[field] = "手动确认"
            self.highlight_layer_id = layer.layer_id
            self.refresh_mapping()
            self.select_mapping_field(field)
            self.status_var.set(f"商品位{slot.slot_index}：{FIELD_LABELS[field]}已绑定到“{self.layer_display(layer)}”")
        else:
            self.highlight_layer_id = layer.layer_id
            self.redraw_preview()

    def load_excel(self):
        if not self.analysis:
            messagebox.showinfo(APP_NAME, "请先分析模板，确定商品位数量。")
            return
        path = Path(self.excel_var.get().strip())
        if not path.exists():
            messagebox.showwarning(APP_NAME, "请选择有效的Excel文件。")
            return
        try:
            self.workbook_data = parse_workbook(path, len(self.analysis.slots))
            page_count = len(self.workbook_data.pages)
            product_count = sum(len(page.products) for page in self.workbook_data.pages)
            warning_text = f"；{len(self.workbook_data.warnings)}条提醒" if self.workbook_data.warnings else ""
            self.excel_info_var.set(f"{self.workbook_data.format_name}：{page_count}张页面、{product_count}条商品数据{warning_text}")
            self.status_var.set("Excel读取完成。可以先生成第1张测试稿。")
            self.refresh_data_preview()
            self.append_activity(f"Excel读取完成：{page_count}张页面、{product_count}条商品数据。")
            self.info_notebook.select(0)
            self.after(100, lambda: self.toggle_source(True))
        except Exception as error:
            self._show_error(f"Excel读取失败：{error}")

    def validate_before_run(self) -> bool:
        if not self.analysis:
            messagebox.showwarning(APP_NAME, "请先分析模板。")
            return False
        if not self.workbook_data:
            self.load_excel()
            if not self.workbook_data:
                return False
        if not self.photoshop_path or not self.photoshop_path.exists():
            messagebox.showerror(APP_NAME, "未检测到Photoshop 2025/2026。")
            return False
        assets = Path(self.assets_var.get().strip())
        if not assets.exists():
            messagebox.showwarning(APP_NAME, "请选择有效的产品素材文件夹。")
            return False
        output = Path(self.output_var.get().strip())
        if not str(output):
            messagebox.showwarning(APP_NAME, "请选择PSD输出文件夹。")
            return False
        missing = validate_assets(self.workbook_data, assets)
        if missing:
            preview = "\n".join(missing[:8])
            messagebox.showerror(APP_NAME, f"有{len(missing)}张产品图不存在：\n\n{preview}")
            return False
        required = ["image", "selling_point", "product_name"]
        missing_bindings = []
        for slot in self.analysis.slots:
            for field in required:
                if not slot.fields.get(field):
                    missing_bindings.append(f"商品位{slot.slot_index}：{FIELD_LABELS[field]}")
        if missing_bindings:
            messagebox.showerror(APP_NAME, "以下必要变量尚未绑定：\n\n" + "\n".join(missing_bindings))
            return False
        return True

    def run_batch(self, test_only: bool):
        if not self.validate_before_run():
            return
        self.save_mapping()
        workbook = self.workbook_data
        if test_only:
            workbook = WorkbookData(source_path=workbook.source_path, format_name=workbook.format_name, pages=workbook.pages[:1], warnings=workbook.warnings)
        jobs = build_jobs(self.analysis, workbook, self.assets_var.get(), self.output_var.get())
        run_dir = self.app_data_dir / "runs" / time.strftime("%Y%m%d-%H%M%S")
        run_dir.mkdir(parents=True, exist_ok=True)
        script_path = run_dir / "batch.jsx"
        self.batch_log_path = run_dir / "batch.log"
        write_batch_script(jobs, script_path, self.batch_log_path)
        try:
            self.batch_process = launch_batch(self.photoshop_path, script_path)
        except Exception as error:
            self._show_error(f"无法启动Photoshop：{error}")
            return
        self.batch_log_position = 0
        self.status_var.set(f"Photoshop正在生成{len(jobs)}张PSD；商品位较多时可能需要数分钟，请勿关闭Photoshop……")
        self.append_activity(f"已提交{len(jobs)}张PSD生成任务，请勿关闭Photoshop。")
        self.info_notebook.select(1)
        self.after(1000, self.poll_batch_log)

    def poll_batch_log(self):
        if self.batch_log_path and self.batch_log_path.exists():
            text = self.batch_log_path.read_text(encoding="utf-8", errors="ignore")
            new_text = text[self.batch_log_position:]
            self.batch_log_position = len(text)
            for line in new_text.splitlines():
                if "PAGE_START" in line:
                    detail = line.split("PAGE_START", 1)[1].strip()
                    self.status_var.set("正在处理：" + detail)
                    self.append_activity("正在处理：" + detail)
                elif "PRODUCT" in line:
                    self.status_var.set("正在替换商品：" + line.split("PRODUCT", 1)[1].strip())
                elif "PAGE_OK" in line:
                    detail = line.split("PAGE_OK", 1)[1].strip()
                    self.status_var.set("已生成：" + detail)
                    self.append_activity("已生成：" + detail)
                elif "PAGE_ERROR" in line:
                    detail = line.split("PAGE_ERROR", 1)[1].strip()
                    self.status_var.set("生成失败：" + detail)
                    self.append_activity("生成失败：" + detail)
            if " DONE" in text or text.rstrip().endswith("DONE"):
                errors = [line for line in text.splitlines() if "PAGE_ERROR" in line]
                if errors:
                    messagebox.showwarning(APP_NAME, f"批量任务完成，但有{len(errors)}张失败。请查看状态日志。")
                else:
                    messagebox.showinfo(APP_NAME, "PSD生成完成。")
                self.status_var.set("任务完成。可以打开输出目录检查PSD。")
                self.append_activity("任务完成，可以打开输出目录检查PSD。")
                return
        self.after(1000, self.poll_batch_log)

    def open_output_folder(self):
        path = Path(self.output_var.get().strip())
        if path.exists():
            os.startfile(path)
        else:
            messagebox.showinfo(APP_NAME, "输出目录尚不存在。")

    def _show_error(self, message: str):
        self.status_var.set(message)
        self.template_info_var.set("发生错误")
        messagebox.showerror(APP_NAME, message)


def main():
    enable_high_dpi()
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--template", default="")
    parser.add_argument("--excel", default="")
    parser.add_argument("--assets", default="")
    parser.add_argument("--output", default="")
    parser.add_argument("--auto-analyze", action="store_true")
    args, _unknown = parser.parse_known_args()
    app = SmartTemplateApp()
    if args.template:
        app.template_var.set(args.template)
    if args.excel:
        app.excel_var.set(args.excel)
    if args.assets:
        app.assets_var.set(args.assets)
    if args.output:
        app.output_var.set(args.output)
    if args.auto_analyze and args.template:
        app.after(300, app.start_analysis)
    app.mainloop()


if __name__ == "__main__":
    main()
