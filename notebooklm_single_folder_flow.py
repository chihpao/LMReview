# notebooklm_single_folder_flow.py
# -*- coding: utf-8 -*-
import os
import sys
import time
import json
import logging
import webbrowser
import re
from logging.handlers import RotatingFileHandler
from dataclasses import dataclass, field
from typing import List, Optional, Tuple
import importlib.util

import customtkinter as ctk
from tkinter import messagebox

from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# ==================== 核心配置 ====================

APP_VERSION = "1.6.0" # Recovery Version

def get_base_path() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

@dataclass
class AppConfig:
    base_path: str = field(default_factory=get_base_path)
    notebook_url: str = "https://notebooklm.google.com/"
    projects: List[str] = field(default_factory=lambda: ["【雲端案】", "【整合案】", "【Trod案】"])
    deliveries: List[str] = field(default_factory=lambda: ["【契約交付】", "【其他交付】"])
    input_folder: str = "input"
    output_folder: str = "output"
    settings_filename: str = "settings.json"
    config_filename: str = "lmreview_config.json"
    tags = ["【標準】", "【範本】", "【待審】"]

    def __post_init__(self):
        self._load_config()

    def _load_config(self):
        path = os.path.join(self.base_path, self.config_filename)
        if os.path.exists(path):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    if data.get("projects"): self.projects = data["projects"]
                    if data.get("deliveries"): self.deliveries = data["deliveries"]
            except Exception as e:
                print(f"Config load error: {e}")

# ==================== 工具與邏輯 ====================

def sanitize_filename(name: str) -> str:
    return re.sub(r'[\\/:*?"<>|]+', "_", name)

def is_skip_file(fn: str) -> bool:
    if fn.lower() in {"thumbs.db", "desktop.ini"}: return True
    return fn.startswith("~$") or fn.startswith(".")

def open_folder(path: str):
    os.makedirs(path, exist_ok=True)
    if os.name == "nt": os.startfile(path)
    else: webbrowser.open(f"file://{os.path.abspath(path)}")

class FileManager:
    def __init__(self, cfg):
        self.cfg = cfg

    def list_files(self, project, delivery):
        path = os.path.join(self.cfg.base_path, project, delivery, self.cfg.input_folder)
        if not os.path.exists(path):
            return [], []
        try:
            files = [f for f in os.listdir(path) if os.path.isfile(os.path.join(path, f)) and not is_skip_file(f)]
            tagged = sorted([f for f in files if any(f.startswith(t) for t in self.cfg.tags)])
            untagged = sorted([f for f in files if not any(f.startswith(t) for t in self.cfg.tags)])
            return tagged, untagged
        except Exception as e:
            print(f"List files error: {e}")
            return [], []

    def tag_file(self, project, delivery, filename, tag):
        base = os.path.join(self.cfg.base_path, project, delivery, self.cfg.input_folder)
        old = os.path.join(base, filename)
        new = os.path.join(base, f"{tag}{filename}")
        try:
            os.rename(old, new)
            return True, new
        except Exception as e:
            return False, str(e)


class WordExporter:
    def __init__(self, logger: logging.Logger):
        self.logger = logger

    def export(self, out_dir: str, target_filename: str, text: str) -> str:
        os.makedirs(out_dir, exist_ok=True)
        base = os.path.splitext(target_filename)[0]
        safe_name = sanitize_filename(base) + ".docx"
        path = os.path.join(out_dir, safe_name)

        try:
            import docx  # type: ignore
        except Exception as e:
            self.logger.error(f"python-docx import error: {e}")
            raise RuntimeError("未安裝 python-docx") from e

        doc = docx.Document()
        for line in text.splitlines():
            doc.add_paragraph(line)
        doc.save(path)
        return path

# ==================== GUI ====================

class NotebookLMSingleFolderApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.cfg = AppConfig()
        self.fm = FileManager(self.cfg)
        self.logger = self._setup_logger()
        self.exporter = WordExporter(self.logger)
        
        self.title(f"LMReview v{APP_VERSION}")
        self.geometry("1280x800")
        ctk.set_appearance_mode("Light")
        
        self.colors = {"panel": "#FFFFFF", "bg": "#F3F4F6", "accent": "#2563EB", "text": "#1F2937"}
        self.colors["file_hover_bg"] = "#F3F4F6"
        self.colors["file_selected_bg"] = "#DBEAFE"
        self.selected_file = None
        self.inbox_buttons = {}
        self._inbox_scrollbar_lock = False
        self._inbox_scrollbar_after_id = None
        self._inbox_unlock_job = None
        self._inbox_scroll_bindings_enabled = False
        self._inbox_config_bindings_enabled = False
        self._inbox_config_unlock_job = None
        self._inbox_generation = 0
        self._inbox_created = True
        
        self._build_ui()
        
        # 延遲刷新，確保 UI 渲染完畢
        self.after(500, self.refresh_all)
        self.bind("<F5>", lambda _e: self.refresh_all())

    def _build_ui(self):
        self.project_var = ctk.StringVar(value=self.cfg.projects[0])
        self.delivery_var = ctk.StringVar(value=self.cfg.deliveries[0])

        # Main Layout
        main = ctk.CTkFrame(self, fg_color=self.colors["bg"])
        main.pack(fill="both", expand=True)
        
        # Left Sidebar
        sidebar = ctk.CTkFrame(main, fg_color="white", width=350, corner_radius=0)
        sidebar.pack(side="left", fill="y", padx=(0, 1))
        sidebar.pack_propagate(False)
        
        # Sidebar Header + Selectors
        ctk.CTkLabel(sidebar, text="專案設定", font=("Microsoft JhengHei UI", 14, "bold"), anchor="w", justify="left").pack(fill="x", padx=20, pady=(20, 8))
        ctk.CTkOptionMenu(sidebar, values=self.cfg.projects, variable=self.project_var, width=200, command=lambda _: self.refresh_all()).pack(fill="x", padx=20, pady=5)
        ctk.CTkOptionMenu(sidebar, values=self.cfg.deliveries, variable=self.delivery_var, width=200, command=lambda _: self.refresh_all()).pack(fill="x", padx=20, pady=(0, 10))

        ctk.CTkLabel(sidebar, text="檔案管理", font=("Microsoft JhengHei UI", 16, "bold"), anchor="w", justify="left").pack(fill="x", padx=20, pady=(5, 10))
        
        ctk.CTkButton(sidebar, text="📂 開啟資料夾", fg_color="#EFF6FF", text_color=self.colors["accent"], hover_color="#DBEAFE", command=self.open_input).pack(fill="x", padx=20, pady=5)
        ctk.CTkButton(sidebar, text="🔄 重新整理", fg_color="#EFF6FF", text_color=self.colors["accent"], hover_color="#DBEAFE", command=self.refresh_all).pack(fill="x", padx=20, pady=5)
        
        # Tabs
        self.tab_var = ctk.StringVar(value="待標記")
        tabs = ctk.CTkSegmentedButton(sidebar, values=["待標記", "已標記"], variable=self.tab_var, command=self.switch_tab)
        tabs.pack(fill="x", padx=20, pady=10)
        
        # File Lists (Container)
        self.list_container = ctk.CTkFrame(sidebar, fg_color="transparent")
        self.list_container.pack(fill="both", expand=True, padx=10, pady=5)
        
        # Inbox View
        self.frame_inbox = ctk.CTkFrame(self.list_container, fg_color="transparent")
        self._build_inbox_scroller()
        
        # Tag Actions (Fixed Bottom of Inbox)
        self.tag_area = ctk.CTkFrame(self.frame_inbox, fg_color="#F9FAFB", corner_radius=10)
        self.tag_area.pack(side="bottom", fill="x", pady=10, padx=5)
        
        self.lbl_selected = ctk.CTkLabel(self.tag_area, text="尚未選擇", text_color="gray")
        self.lbl_selected.pack(pady=5)
        
        btn_row = ctk.CTkFrame(self.tag_area, fg_color="transparent")
        btn_row.pack(fill="x", pady=5)
        btn_row.grid_columnconfigure(0, weight=1)
        btn_row.grid_columnconfigure(1, weight=1)
        btn_row.grid_columnconfigure(2, weight=1)
        
        tag_colors = {
            "【標準】": "#2563EB",
            "【範本】": "#16A34A",
            "【待審】": "#F59E0B",
        }
        for i, t in enumerate(self.cfg.tags):
            ctk.CTkButton(
                btn_row,
                text=t,
                height=32,
                fg_color=tag_colors.get(t, self.colors["accent"]),
                hover_color=tag_colors.get(t, self.colors["accent"]),
                text_color="white",
                command=lambda tag=t: self.do_tag(tag),
            ).grid(row=0, column=i, sticky="ew", padx=2)

        # Tagged View
        self.frame_tagged = ctk.CTkFrame(self.list_container, fg_color="transparent")
        self.txt_tagged = ctk.CTkTextbox(self.frame_tagged, fg_color="transparent")
        try:
            self.txt_tagged.configure(wrap="word")
        except Exception:
            try:
                self.txt_tagged._textbox.configure(wrap="word")
            except Exception:
                pass
        # Force-hide vertical scrollbar in tagged view
        try:
            self.txt_tagged._scrollbar.pack_forget()
        except Exception:
            pass
        self.txt_tagged.pack(fill="both", expand=True)

        self.frame_inbox.pack(fill="both", expand=True) # Default show inbox

        # Right Workspace
        workspace = ctk.CTkFrame(main, fg_color="transparent")
        workspace.pack(side="left", fill="both", expand=True, padx=20, pady=20)
        
        # Prompt Area
        ws_top = ctk.CTkFrame(workspace, fg_color="white", corner_radius=10)
        ws_top.pack(fill="both", expand=True, pady=(0, 10))
        
        ctk.CTkLabel(ws_top, text="生成提示詞", font=("Microsoft JhengHei UI", 16, "bold")).pack(anchor="w", padx=20, pady=(15, 5))
        
        ctrl_row = ctk.CTkFrame(ws_top, fg_color="transparent")
        ctrl_row.pack(fill="x", padx=20, pady=5)

        ctrl_height = 36
        self.review_var = ctk.StringVar(value="(無)")
        ctrl_row.grid_columnconfigure(0, weight=1)
        ctrl_row.grid_rowconfigure(0, weight=0)
        ctrl_row.grid_rowconfigure(1, weight=0)
        self.review_menu = ctk.CTkOptionMenu(ctrl_row, variable=self.review_var, values=["(無)"], height=ctrl_height)
        self.review_menu.grid(row=0, column=0, sticky="ew")
        
        actions_row = ctk.CTkFrame(ctrl_row, fg_color="transparent")
        actions_row.grid(row=0, column=1, sticky="e", padx=(10, 0))
        actions_row.grid_columnconfigure(0, weight=0)
        actions_row.grid_columnconfigure(1, weight=0)

        ctk.CTkButton(actions_row, text="✨ 生成", width=80, height=ctrl_height, command=self.gen_prompt).grid(row=0, column=0, padx=(0, 10))
        ctk.CTkButton(actions_row, text="📋 複製", width=60, height=ctrl_height, fg_color="gray", command=self.copy_prompt).grid(row=0, column=1)
        self.lbl_copy_status = ctk.CTkLabel(ctrl_row, text="", text_color="gray", font=("Microsoft JhengHei UI", 11))
        self.lbl_copy_status.grid(row=1, column=1, pady=(3, 0), sticky="e")
        
        self.txt_prompt = ctk.CTkTextbox(ws_top)
        self.txt_prompt.pack(fill="both", expand=True, padx=20, pady=10)

        # Reply Area
        ws_bot = ctk.CTkFrame(workspace, fg_color="white", corner_radius=10)
        ws_bot.pack(fill="both", expand=True)
        
        # 1. Header (Top)
        ctk.CTkLabel(ws_bot, text="AI 回覆處理", font=("Microsoft JhengHei UI", 16, "bold")).pack(anchor="w", padx=20, pady=(15, 5))
        
        # 2. Controls (Bottom First - Ensure Visibility)
        bot_ctrl = ctk.CTkFrame(ws_bot, fg_color="transparent")
        bot_ctrl.pack(side="bottom", fill="x", padx=20, pady=(0, 15))
        
        ctk.CTkButton(bot_ctrl, text="貼上", width=80, fg_color="gray", command=self.paste_reply).pack(side="left")
        ctk.CTkButton(bot_ctrl, text="📄 輸出 Word", width=120, command=self.export_word).pack(side="right")
        ctk.CTkButton(bot_ctrl, text="開啟 output", width=100, fg_color="transparent", text_color="gray", command=self.open_output).pack(side="right", padx=10)

        # 3. Text Area (Fill Remaining)
        self.txt_reply = ctk.CTkTextbox(ws_bot)
        self.txt_reply.pack(side="top", fill="both", expand=True, padx=20, pady=10)

        # Status Bar
        self.status = ctk.CTkLabel(self, text="Ready", height=25, fg_color="#E5E7EB", anchor="w")
        self.status.pack(side="bottom", fill="x")

    def _setup_logger(self) -> logging.Logger:
        logger = logging.getLogger("lmreview")
        if logger.handlers:
            return logger
        logger.setLevel(logging.INFO)
        logs_dir = os.path.join(self.cfg.base_path, "logs")
        os.makedirs(logs_dir, exist_ok=True)
        log_path = os.path.join(logs_dir, "lmreview.log")
        handler = RotatingFileHandler(log_path, maxBytes=512_000, backupCount=3, encoding="utf-8")
        formatter = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")
        handler.setFormatter(formatter)
        logger.addHandler(handler)
        return logger

    # --- Actions ---

    def switch_tab(self, val):
        self._inbox_generation += 1
        self._lock_inbox_scrollbar()
        if val == "待標記":
            self.frame_tagged.pack_forget()
            self.frame_inbox.pack(fill="both", expand=True)
            if not self._inbox_created:
                self._build_inbox_scroller()
                self.refresh_all()
            self._schedule_inbox_unlock()
            self._set_inbox_scroll_bindings(True)
            self._schedule_inbox_config_unlock()
        else:
            if self._inbox_created:
                self._destroy_inbox_scroller()
            self.frame_inbox.pack_forget()
            self.frame_tagged.pack(fill="both", expand=True)

    def open_input(self):
        p, d = self.project_var.get(), self.delivery_var.get()
        path = os.path.join(self.cfg.base_path, p, d, self.cfg.input_folder)
        open_folder(path)

    def open_output(self):
        p, d = self.project_var.get(), self.delivery_var.get()
        path = os.path.join(self.cfg.base_path, p, d, self.cfg.output_folder)
        open_folder(path)

    def refresh_all(self):
        p, d = self.project_var.get(), self.delivery_var.get()
        print(f"[UI] Refreshing: {p} / {d}") # Debug to console
        tagged, untagged = self.fm.list_files(p, d)
        if not self._inbox_created:
            # When inbox UI is not present (e.g. in tagged view), skip rebuilding list
            self.txt_tagged.delete("1.0", "end")
            self.txt_tagged.insert("1.0", "\n".join(tagged))
            review_files = [f for f in tagged if f.startswith("【待審】")]
            self.review_menu.configure(values=review_files if review_files else ["(無)"])
            if review_files: self.review_var.set(review_files[0])
            else: self.review_var.set("(無)")
            self.status.configure(text=f"  Path: {p}/{d} | 待標記: {len(untagged)}")
            self.select_file(None)
            return
        
        # Refresh Untagged
        for w in self.inbox_content.winfo_children(): w.destroy()
        self.inbox_buttons = {}
        if not untagged:
            ctk.CTkLabel(self.inbox_content, text="(無檔案)", text_color="gray").pack(pady=20)
        else:
            for f in untagged:
                btn = ctk.CTkButton(
                    self.inbox_content,
                    text=f"📄 {f}",
                    fg_color="transparent",
                    hover_color=self.colors["file_hover_bg"],
                    text_color="black",
                    anchor="w",
                    command=lambda x=f: self.select_file(x),
                )
                btn.pack(fill="x", pady=2)
                self.inbox_buttons[f] = btn
        
        # Refresh Tagged
        self.txt_tagged.delete("1.0", "end")
        self.txt_tagged.insert("1.0", "\n".join(tagged))
        
        # Refresh Dropdown
        review_files = [f for f in tagged if f.startswith("【待審】")]
        self.review_menu.configure(values=review_files if review_files else ["(無)"])
        if review_files: self.review_var.set(review_files[0])
        else: self.review_var.set("(無)")

        self.status.configure(text=f"  Path: {p}/{d} | 待標記: {len(untagged)}")
        self.select_file(None)
        self.after(0, self._update_inbox_scrollbar)

    def select_file(self, f):
        self.selected_file = f
        self.lbl_selected.configure(text=f"已選: {f}" if f else "尚未選擇")
        self._update_inbox_selection()

    def _update_inbox_selection(self):
        for name, btn in self.inbox_buttons.items():
            if name == self.selected_file:
                btn.configure(fg_color=self.colors["file_selected_bg"])
            else:
                btn.configure(fg_color="transparent")

    def _build_inbox_scroller(self):
        self.scroll_inbox = ctk.CTkFrame(self.frame_inbox, fg_color="transparent")
        if hasattr(self, "tag_area"):
            self.scroll_inbox.pack(side="top", fill="both", expand=True, before=self.tag_area)
        else:
            self.scroll_inbox.pack(side="top", fill="both", expand=True)
        self.inbox_canvas = ctk.CTkCanvas(self.scroll_inbox, bg="white", highlightthickness=0)
        self.inbox_scrollbar = None
        self.inbox_canvas.pack(side="left", fill="both", expand=True)
        self._inbox_scrollbar_cfg = {}

        self.inbox_content = ctk.CTkFrame(self.inbox_canvas, fg_color="transparent")
        self.inbox_window = self.inbox_canvas.create_window((0, 0), window=self.inbox_content, anchor="nw")
        self._set_inbox_config_bindings(True)
        self._bind_inbox_mousewheel()
        self._inbox_created = True

    def _destroy_inbox_scroller(self):
        self._set_inbox_scroll_bindings(False)
        self._set_inbox_config_bindings(False)
        if self.inbox_scrollbar:
            try:
                if self.inbox_scrollbar.winfo_exists():
                    self.inbox_scrollbar.destroy()
            except Exception:
                pass
        self.inbox_scrollbar = None
        self._inbox_scrollbar_cfg = {}
        try:
            if self.inbox_canvas.winfo_exists():
                self.inbox_canvas.destroy()
        except Exception:
            pass
        try:
            if self.scroll_inbox.winfo_exists():
                self.scroll_inbox.destroy()
        except Exception:
            pass
        self._inbox_created = False

    def _on_inbox_canvas_configure(self, event):
        try:
            if self._inbox_scrollbar_lock or not self._inbox_config_bindings_enabled:
                return
            self.inbox_canvas.itemconfigure(self.inbox_window, width=event.width)
            self._update_inbox_scrollbar()
        except Exception:
            pass

    def _bind_inbox_mousewheel(self):
        self._set_inbox_scroll_bindings(True)

    def _set_inbox_scroll_bindings(self, enabled: bool):
        if enabled and not self._inbox_scroll_bindings_enabled:
            self.inbox_canvas.bind("<MouseWheel>", self._on_inbox_mousewheel)
            self.inbox_content.bind("<MouseWheel>", self._on_inbox_mousewheel)
            self._inbox_scroll_bindings_enabled = True
        elif not enabled and self._inbox_scroll_bindings_enabled:
            self.inbox_canvas.unbind("<MouseWheel>")
            self.inbox_content.unbind("<MouseWheel>")
            self._inbox_scroll_bindings_enabled = False

    def _set_inbox_config_bindings(self, enabled: bool):
        if enabled and not self._inbox_config_bindings_enabled:
            self.inbox_content.bind("<Configure>", lambda _e: self._update_inbox_scrollbar())
            self.inbox_canvas.bind("<Configure>", self._on_inbox_canvas_configure)
            self._inbox_config_bindings_enabled = True
        elif not enabled and self._inbox_config_bindings_enabled:
            self.inbox_content.unbind("<Configure>")
            self.inbox_canvas.unbind("<Configure>")
            self._inbox_config_bindings_enabled = False

    def _on_inbox_mousewheel(self, event):
        if self._inbox_scrollbar_lock or self.tab_var.get() != "待標記":
            return
        if os.name == "nt":
            self.inbox_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        else:
            self.inbox_canvas.yview_scroll(int(-1 * event.delta), "units")

    def _update_inbox_scrollbar(self):
        try:
            if self._inbox_scrollbar_lock:
                return
            if self.tab_var.get() != "待標記":
                return
            gen = self._inbox_generation
            if self._inbox_scrollbar_after_id:
                self.after_cancel(self._inbox_scrollbar_after_id)
                self._inbox_scrollbar_after_id = None
            self._inbox_scrollbar_after_id = self.after(100, lambda g=gen: self._apply_inbox_scrollbar(g))
        except Exception:
            pass

    def _lock_inbox_scrollbar(self):
        self._inbox_scrollbar_lock = True
        if self._inbox_scrollbar_after_id:
            self.after_cancel(self._inbox_scrollbar_after_id)
            self._inbox_scrollbar_after_id = None
        if self._inbox_unlock_job:
            self.after_cancel(self._inbox_unlock_job)
            self._inbox_unlock_job = None
        if self._inbox_config_unlock_job:
            self.after_cancel(self._inbox_config_unlock_job)
            self._inbox_config_unlock_job = None
        self._set_inbox_config_bindings(False)

    def _schedule_inbox_unlock(self):
        if self._inbox_unlock_job:
            self.after_cancel(self._inbox_unlock_job)
        self._inbox_unlock_job = self.after(200, self._unlock_inbox_scrollbar)

    def _schedule_inbox_config_unlock(self):
        if self._inbox_config_unlock_job:
            self.after_cancel(self._inbox_config_unlock_job)
        self._inbox_config_unlock_job = self.after(300, self._unlock_inbox_config_bindings)

    def _unlock_inbox_config_bindings(self):
        self._inbox_config_unlock_job = None
        if self.tab_var.get() != "待標記":
            return
        self._set_inbox_config_bindings(True)

    def _unlock_inbox_scrollbar(self):
        self._inbox_unlock_job = None
        if self.tab_var.get() != "待標記":
            return
        self._inbox_scrollbar_lock = False
        self._update_inbox_scrollbar()

    def _apply_inbox_scrollbar(self, gen: int):
        try:
            if gen != self._inbox_generation:
                return
            if self._inbox_scrollbar_lock or self.tab_var.get() != "待標記":
                if self.inbox_scrollbar:
                    try:
                        if self.inbox_scrollbar.winfo_exists():
                            self.inbox_scrollbar.destroy()
                    except Exception:
                        pass
                    self.inbox_scrollbar = None
                    self._inbox_scrollbar_cfg = {}
                self.inbox_canvas.configure(yscrollcommand=None)
                return

            self.inbox_canvas.update_idletasks()
            content_h = self.inbox_content.winfo_reqheight()
            view_h = self.inbox_canvas.winfo_height()
            if view_h <= 10:
                self._inbox_scrollbar_after_id = self.after(100, self._apply_inbox_scrollbar)
                return
            need_scroll = content_h > (view_h + 4)
            self.inbox_canvas.configure(scrollregion=(0, 0, self.inbox_canvas.winfo_width(), content_h))

            if need_scroll:
                if not self.inbox_scrollbar:
                    self.inbox_scrollbar = ctk.CTkScrollbar(self.scroll_inbox, command=self.inbox_canvas.yview)
                    self.inbox_canvas.configure(yscrollcommand=self.inbox_scrollbar.set)
                    self._inbox_scrollbar_cfg = {
                        "width": self.inbox_scrollbar.cget("width"),
                        "fg_color": self.inbox_scrollbar.cget("fg_color"),
                    }
                if not self.inbox_scrollbar.winfo_ismapped():
                    self.inbox_scrollbar.pack(side="right", fill="y")
            else:
                if self.inbox_scrollbar:
                    try:
                        if self.inbox_scrollbar.winfo_exists():
                            self.inbox_scrollbar.destroy()
                    except Exception:
                        pass
                    self.inbox_scrollbar = None
                    self._inbox_scrollbar_cfg = {}
                self.inbox_canvas.configure(yscrollcommand=None)
                self.inbox_canvas.yview_moveto(0)
            self._sync_inbox_width()
        except Exception:
            pass

    def _sync_inbox_width(self):
        try:
            self.inbox_canvas.update_idletasks()
            self.inbox_canvas.itemconfigure(self.inbox_window, width=self.inbox_canvas.winfo_width())
        except Exception:
            pass

    def do_tag(self, tag):
        if not self.selected_file: return
        p, d = self.project_var.get(), self.delivery_var.get()
        ok, msg = self.fm.tag_file(p, d, self.selected_file, tag)
        if ok: self.refresh_all()
        else: messagebox.showerror("Error", msg)

    def gen_prompt(self):
        p, d = self.project_var.get(), self.delivery_var.get()
        tagged, _ = self.fm.list_files(p, d)
        std = [f for f in tagged if f.startswith("【標準】")]
        tpl = [f for f in tagged if f.startswith("【範本】")]
        tgt = self.review_var.get()
        if tgt == "(無)": return
        
        txt = f"開始審查：{tgt}"
        self.txt_prompt.delete("1.0", "end")
        self.txt_prompt.insert("1.0", txt)

    def copy_prompt(self):
        self.clipboard_clear()
        self.clipboard_append(self.txt_prompt.get("1.0", "end"))
        self.lbl_copy_status.configure(text="複製成功")
        self.after(1200, lambda: self.lbl_copy_status.configure(text=""))

    def paste_reply(self):
        try: t = self.clipboard_get()
        except: t = ""
        self.txt_reply.delete("1.0", "end")
        self.txt_reply.insert("1.0", t)

    def export_word(self):
        txt = self.txt_reply.get("1.0", "end").strip()
        if not txt:
            messagebox.showwarning("提醒", "回覆內容是空的，請先貼上 AI 的回覆")
            return
            
        p, d = self.project_var.get(), self.delivery_var.get()
        tgt = self.review_var.get()
        if tgt == "(無)":
            messagebox.showwarning("提醒", "請先選擇對應的【待審】檔案，以便命名 Word 報告")
            return

        try:
            out_dir = os.path.join(self.cfg.base_path, p, d, self.cfg.output_folder)
            path = self.exporter.export(out_dir, tgt, txt)
            
            # Success feedback
            self.status.configure(text=f"✅ Word 已輸出: {os.path.basename(path)}")
            # Auto open folder
            open_folder(os.path.dirname(path))
            
        except Exception as e:
            self.logger.error(f"Export Error: {e}")
            messagebox.showerror("輸出錯誤", f"Word 產出失敗：\n{e}\n\n請確認是否已安裝 python-docx")

    def _on_close(self):
        self.destroy()

if __name__ == "__main__":
    app = NotebookLMSingleFolderApp()
    app.mainloop()
