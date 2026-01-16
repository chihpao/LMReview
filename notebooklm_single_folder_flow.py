# notebooklm_single_folder_flow.py
# -*- coding: utf-8 -*-
import os
import re
import sys
import time
import logging
import webbrowser
from dataclasses import dataclass, field
from typing import List, Optional, Tuple

import customtkinter as ctk
from tkinter import messagebox, TclError

from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# ==================== 基本設定 ====================

APP_VERSION = "1.2.0"

def get_base_path() -> str:
    """取得程式基底路徑（支援 PyInstaller）"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

@dataclass
class AppConfig:
    base_path: str = field(default_factory=get_base_path)
    notebook_url: str = "https://notebooklm.google.com/"

    projects: Optional[List[str]] = None
    deliveries: Optional[List[str]] = None

    input_folder: str = "input"
    output_folder: str = "output"

    tags = ["【標準】", "【範本】", "【待審】"]

    def __post_init__(self):
        if self.projects is None:
            self.projects = ["【雲端案】", "【整合案】", "【Trod案】"]
        if self.deliveries is None:
            self.deliveries = ["【契約交付】", "【其他交付】"]

# ==================== 工具 ====================

def sanitize_filename(name: str) -> str:
    return re.sub(r'[\\/:*?"<>|]+', "_", name)

def shorten_path(path: str, max_len: int = 60) -> str:
    if len(path) <= max_len:
        return path
    head = max_len // 2 - 2
    tail = max_len - head - 3
    return f"{path[:head]}...{path[-tail:]}"

def now_ts() -> str:
    return time.strftime("%Y%m%d_%H%M%S")

def is_skip_file(fn: str) -> bool:
    return fn.startswith("~$") or fn.startswith(".") or fn.endswith(".tmp")

def open_folder(path: str):
    os.makedirs(path, exist_ok=True)
    if os.name == 'nt':
        os.startfile(path)
    else:
        webbrowser.open(f'file://{path}')

# ==================== 檔案管理 ====================

class FileManager:
    def __init__(self, cfg: AppConfig, logger: logging.Logger):
        self.cfg = cfg
        self.logger = logger

    def project_root(self, project: str, delivery: str) -> str:
        return os.path.join(self.cfg.base_path, project, delivery)

    def input_dir(self, project: str, delivery: str) -> str:
        return os.path.join(self.project_root(project, delivery), self.cfg.input_folder)

    def output_dir(self, project: str, delivery: str) -> str:
        return os.path.join(self.project_root(project, delivery), self.cfg.output_folder)

    def ensure_structure(self):
        for p in self.cfg.projects:
            for d in self.cfg.deliveries:
                os.makedirs(self.input_dir(p, d), exist_ok=True)
                os.makedirs(self.output_dir(p, d), exist_ok=True)

    def list_input_files(self, project: str, delivery: str) -> Tuple[List[str], List[str]]:
        """返回 (已標記檔案列表, 未標記檔案列表)"""
        folder = self.input_dir(project, delivery)
        if not os.path.exists(folder):
            return [], []
        
        all_files = [
            f for f in os.listdir(folder)
            if os.path.isfile(os.path.join(folder, f)) and not is_skip_file(f)
        ]
        
        tagged = [f for f in all_files if any(f.startswith(t) for t in self.cfg.tags)]
        untagged = [f for f in all_files if not any(f.startswith(t) for t in self.cfg.tags)]
        
        return sorted(tagged), sorted(untagged)

    def tag_file(self, project: str, delivery: str, filename: str, tag: str) -> Tuple[bool, str]:
        """為檔案加上標籤"""
        directory = self.input_dir(project, delivery)
        old_path = os.path.join(directory, filename)
        new_filename = f"{tag}{filename}"
        new_path = os.path.join(directory, new_filename)
        
        try:
            if not os.path.exists(old_path):
                return False, "檔案不存在"
            
            if os.path.exists(new_path):
                return False, "目標檔名已存在"
            
            # 等待檔案穩定
            for _ in range(5):
                try:
                    size1 = os.path.getsize(old_path)
                    time.sleep(0.1)
                    size2 = os.path.getsize(old_path)
                    if size1 == size2:
                        break
                except:
                    time.sleep(0.1)
            
            os.rename(old_path, new_path)
            self.logger.info(f"✓ 標記完成：{filename} → {new_filename}")
            return True, new_filename
            
        except PermissionError:
            return False, "檔案被佔用，請關閉後再試"
        except Exception as e:
            self.logger.error(f"標記失敗：{e}")
            return False, str(e)

# ==================== Word 輸出 ====================

class WordExporter:
    def __init__(self, logger: logging.Logger):
        self.logger = logger

    def export(self, output_dir: str, source_filename: str, content: str) -> str:
        try:
            from docx import Document
        except ImportError:
            self.logger.error("缺少 python-docx，執行：pip install python-docx")
            raise

        os.makedirs(output_dir, exist_ok=True)

        safe = sanitize_filename(source_filename)
        path = os.path.join(output_dir, f"Review_{safe}_{now_ts()}.docx")

        doc = Document()
        
        # 標題
        heading = doc.add_heading(f"{source_filename} 審查結果", level=1)
        
        # 內容
        for line in content.splitlines():
            stripped = line.strip()
            if not stripped:
                continue
            if stripped.startswith(("-", "•", "●")):
                doc.add_paragraph(stripped.lstrip("-•● ").strip(), style="List Bullet")
            else:
                doc.add_paragraph(stripped)

        doc.save(path)
        self.logger.info(f"Word 已輸出：{path}")
        return path

# ==================== Watchdog ====================

class AutoTagHandler(FileSystemEventHandler):
    def __init__(self, app):
        super().__init__()
        self.app = app
        self.processing = set()

    def on_created(self, event):
        if event.is_directory:
            return
        
        filepath = event.src_path
        filename = os.path.basename(filepath)
        
        if is_skip_file(filename) or filepath in self.processing:
            return
        
        self.processing.add(filepath)
        # 延遲處理，確保檔案寫入完成
        self.app.after(800, lambda: self.handle_file(filepath))

    def on_modified(self, event):
        if event.is_directory:
            return
        filename = os.path.basename(event.src_path)
        if is_skip_file(filename):
            return
        self.app.schedule_refresh(300)

    def on_moved(self, event):
        if event.is_directory:
            return
        self.app.schedule_refresh(300)

    def on_deleted(self, event):
        if event.is_directory:
            return
        self.app.schedule_refresh(300)

    def handle_file(self, filepath: str):
        try:
            if not os.path.exists(filepath):
                self.processing.discard(filepath)
                return
            
            # 通知用戶有新檔案
            self.app.show_notification("📥 偵測到新檔案，請標記")
            self.app.schedule_refresh(200)
            
        except Exception as e:
            self.app.logger.error(f"處理檔案錯誤：{e}")
        finally:
            self.processing.discard(filepath)

# ==================== GUI ====================

class NotebookLMSingleFolderApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.cfg = AppConfig()
        self.logger = self._setup_logger()
        self.fm = FileManager(self.cfg, self.logger)
        self.exporter = WordExporter(self.logger)

        self._ensure_structure()

        self.title(f"LMReview 文件審查工作流程 v{APP_VERSION}")
        self.geometry("1440x900")
        ctk.set_appearance_mode("Light")

        self.colors = {
            "bg": "#f2f2f7",
            "panel": "#ffffff",
            "sidebar": "#f6f6f8",
            "text": "#1c1c1e",
            "muted": "#6e6e73",
            "border": "#d1d1d6",
            "accent": "#007aff",
            "accent_hover": "#0060df",
            "success": "#34c759",
            "danger": "#ff3b30"
        }
        self.fonts = {
            "title": ctk.CTkFont(family="Microsoft JhengHei UI", size=22, weight="bold"),
            "section": ctk.CTkFont(family="Microsoft JhengHei UI", size=16, weight="bold"),
            "body": ctk.CTkFont(family="Microsoft JhengHei UI", size=13),
            "small": ctk.CTkFont(family="Microsoft JhengHei UI", size=12)
        }

        self.observer = None
        self._refresh_job = None
        self._clipboard_job = None
        self._last_clipboard = None
        self._clipboard_poll_ms = 700

        self._build_ui()
        self.refresh_all()
        self._start_watchdog()

    def _setup_logger(self):
        logger = logging.getLogger("NotebookLM")
        logger.setLevel(logging.INFO)
        if not logger.handlers:
            self._log_formatter = logging.Formatter("%(asctime)s - %(message)s")
            h = logging.StreamHandler()
            h.setFormatter(self._log_formatter)
            logger.addHandler(h)
        self._add_file_handler(logger)
        return logger

    def _add_file_handler(self, logger: logging.Logger):
        if any(isinstance(h, logging.FileHandler) for h in logger.handlers):
            return
        formatter = getattr(self, "_log_formatter", logging.Formatter("%(asctime)s - %(message)s"))
        try:
            log_dir = os.path.join(self.cfg.base_path, "logs")
            os.makedirs(log_dir, exist_ok=True)
            log_file = os.path.join(log_dir, f"notebooklm_{time.strftime('%Y%m%d')}.log")
            fh = logging.FileHandler(log_file, encoding="utf-8")
            fh.setFormatter(formatter)
            logger.addHandler(fh)
        except Exception:
            logger.warning("無法建立日誌檔案，將只輸出到主控台")

    def _ensure_structure(self):
        try:
            self.fm.ensure_structure()
        except PermissionError:
            fallback = os.path.join(os.path.expanduser("~"), "LMReview_Review")
            self.logger.warning("原始路徑無法寫入，改用使用者資料夾：%s", fallback)
            self.cfg.base_path = fallback
            self.fm = FileManager(self.cfg, self.logger)
            self.fm.ensure_structure()
            self._add_file_handler(self.logger)
            messagebox.showwarning("資料夾權限", f"原始路徑無法寫入，已改用：\n{fallback}")

    def _build_ui(self):
        # ========== 頂部導航列 ==========
        top_bar = ctk.CTkFrame(self, height=64, fg_color=self.colors["panel"])
        top_bar.pack(fill="x", padx=0, pady=0)
        top_bar.pack_propagate(False)

        title_label = ctk.CTkLabel(
            top_bar,
            text="LMReview",
            font=self.fonts["title"],
            text_color=self.colors["text"]
        )
        title_label.pack(side="left", padx=24, pady=16)

        project_frame = ctk.CTkFrame(top_bar, fg_color="transparent")
        project_frame.pack(side="right", padx=24, pady=12)

        ctk.CTkLabel(
            project_frame,
            text="專案",
            font=self.fonts["small"],
            text_color=self.colors["muted"]
        ).pack(side="left", padx=(0, 8))

        menu_kwargs = {
            "fg_color": self.colors["panel"],
            "button_color": self.colors["border"],
            "button_hover_color": self.colors["border"],
            "text_color": self.colors["text"],
            "dropdown_fg_color": self.colors["panel"],
            "dropdown_hover_color": self.colors["sidebar"],
            "dropdown_text_color": self.colors["text"]
        }

        self.combo_project = ctk.CTkOptionMenu(
            project_frame,
            values=self.cfg.projects,
            command=lambda _: self.on_selection_change(),
            width=140,
            height=32,
            font=self.fonts["body"],
            **menu_kwargs
        )
        self.combo_project.pack(side="left", padx=5)

        ctk.CTkLabel(
            project_frame,
            text="交付",
            font=self.fonts["small"],
            text_color=self.colors["muted"]
        ).pack(side="left", padx=(12, 8))

        self.combo_delivery = ctk.CTkOptionMenu(
            project_frame,
            values=self.cfg.deliveries,
            command=lambda _: self.on_selection_change(),
            width=140,
            height=32,
            font=self.fonts["body"],
            **menu_kwargs
        )
        self.combo_delivery.pack(side="left", padx=5)

        help_btn = ctk.CTkButton(
            project_frame,
            text="使用說明",
            command=self.show_help,
            width=90,
            height=32,
            font=self.fonts["small"],
            fg_color=self.colors["accent"],
            hover_color=self.colors["accent_hover"],
            text_color="#ffffff",
            corner_radius=16
        )
        help_btn.pack(side="left", padx=(12, 0))

        divider = ctk.CTkFrame(self, height=1, fg_color=self.colors["border"])
        divider.pack(fill="x")

        # ========== 主要內容區域 ==========
        content = ctk.CTkFrame(self, fg_color=self.colors["bg"])
        content.pack(fill="both", expand=True, padx=0, pady=0)

        content.grid_rowconfigure(0, weight=1)
        content.grid_rowconfigure(1, weight=0)
        content.grid_columnconfigure(0, weight=0, minsize=320)
        content.grid_columnconfigure(1, weight=1)

        left_panel = ctk.CTkFrame(content, fg_color="transparent")
        left_panel.grid(row=0, column=0, sticky="nsew", padx=(20, 10), pady=20)

        right_panel = ctk.CTkFrame(content, fg_color="transparent")
        right_panel.grid(row=0, column=1, sticky="nsew", padx=(10, 20), pady=20)

        # ========== 左側：檔案管理 ==========
        self._build_file_panel(left_panel)

        # ========== 右側：工作流程 ==========
        self._build_workflow_panel(right_panel)

        self.notification = ctk.CTkLabel(
            content,
            text="",
            font=self.fonts["small"],
            text_color=self.colors["accent"],
            fg_color=self.colors["panel"],
            corner_radius=10,
            height=0,
            anchor="w"
        )
        self.notification.grid(row=1, column=0, columnspan=2, sticky="ew", padx=20, pady=(0, 12))

        # ========== 底部狀態列 ==========
        status_bar = ctk.CTkFrame(self, height=32, fg_color=self.colors["sidebar"])
        status_bar.pack(fill="x", side="bottom")
        status_bar.pack_propagate(False)

        self.status_left = ctk.CTkLabel(
            status_bar,
            text="",
            font=self.fonts["small"],
            text_color=self.colors["muted"],
            anchor="w"
        )
        self.status_left.pack(side="left", padx=15)

        self.status_right = ctk.CTkLabel(
            status_bar,
            text="",
            font=self.fonts["small"],
            text_color=self.colors["muted"],
            anchor="e"
        )
        self.status_right.pack(side="right", padx=15)

    def _build_file_panel(self, parent):
        """左側：檔案標記區"""
        sidebar = ctk.CTkFrame(parent, fg_color=self.colors["sidebar"], corner_radius=18)
        sidebar.pack(fill="both", expand=True)

        header = ctk.CTkFrame(sidebar, fg_color="transparent")
        header.pack(fill="x", padx=16, pady=(16, 8))
        header.grid_columnconfigure(0, weight=0)
        header.grid_columnconfigure(1, weight=1)
        header.grid_columnconfigure(2, weight=0)

        ctk.CTkLabel(
            header,
            text="檔案",
            font=self.fonts["section"],
            text_color=self.colors["text"]
        ).grid(row=0, column=0, sticky="w")

        self.file_tab_var = ctk.StringVar(value="待標記")
        self.file_tabs = ctk.CTkSegmentedButton(
            header,
            values=["待標記", "已標記"],
            variable=self.file_tab_var,
            command=self._on_file_tab_change,
            height=30,
            font=self.fonts["small"],
            fg_color=self.colors["sidebar"],
            selected_color=self.colors["panel"],
            selected_hover_color=self.colors["panel"],
            unselected_color=self.colors["sidebar"],
            unselected_hover_color=self.colors["border"],
            text_color=self.colors["text"],
            border_width=2
        )
        self.file_tabs.grid(row=0, column=1, sticky="ew", padx=12)

        refresh_btn = ctk.CTkButton(
            header,
            text="重新整理",
            command=self.refresh_all,
            height=28,
            width=72,
            font=self.fonts["small"],
            fg_color=self.colors["panel"],
            hover_color=self.colors["border"],
            text_color=self.colors["text"],
            border_width=1,
            border_color=self.colors["border"],
            corner_radius=14
        )
        refresh_btn.grid(row=0, column=2, sticky="e")

        body = ctk.CTkFrame(sidebar, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        self.file_tab_inbox = ctk.CTkFrame(body, fg_color="transparent")
        self.file_tab_tagged = ctk.CTkFrame(body, fg_color="transparent")

        self.untagged_header_label = ctk.CTkLabel(
            self.file_tab_inbox,
            text="待標記檔案",
            font=self.fonts["small"],
            text_color=self.colors["muted"]
        )
        self.untagged_header_label.pack(anchor="w", padx=8, pady=(10, 6))

        open_btn = ctk.CTkButton(
            self.file_tab_inbox,
            text="開啟 input",
            command=self.open_input,
            height=32,
            font=self.fonts["small"],
            fg_color=self.colors["panel"],
            hover_color=self.colors["border"],
            text_color=self.colors["text"],
            border_width=1,
            border_color=self.colors["border"],
            corner_radius=14
        )
        open_btn.pack(fill="x", padx=8, pady=(0, 8))

        tip = ctk.CTkLabel(
            self.file_tab_inbox,
            text="把檔案放入 input 資料夾，在這裡快速標記",
            font=self.fonts["small"],
            text_color=self.colors["muted"]
        )
        tip.pack(anchor="w", padx=8, pady=(0, 10))

        list_frame = ctk.CTkFrame(
            self.file_tab_inbox,
            fg_color=self.colors["panel"],
            corner_radius=14,
            border_width=1,
            border_color=self.colors["border"]
        )
        list_frame.pack(fill="both", expand=True, padx=8, pady=(0, 10))

        self.untagged_container = ctk.CTkScrollableFrame(
            list_frame,
            fg_color="transparent"
        )
        self.untagged_container.pack(fill="both", expand=True, padx=8, pady=8)

        self._build_tagged_panel(self.file_tab_tagged)
        self._show_file_tab(self.file_tab_var.get())

    def _build_workflow_panel(self, parent):
        """中間：審查工作流程"""
        parent.grid_rowconfigure(0, weight=2)
        parent.grid_rowconfigure(1, weight=3)
        parent.grid_columnconfigure(0, weight=1)

        prompt_card, _ = self._create_step_card(
            parent,
            "提示詞",
            "選擇待審檔案並生成提示詞"
        )
        prompt_card.grid(row=0, column=0, sticky="nsew", pady=(0, 12))

        reply_card, reply_header = self._create_step_card(
            parent,
            "AI 回覆",
            "貼上或監聽剪貼簿後輸出"
        )
        reply_card.grid(row=1, column=0, sticky="nsew")

        menu_kwargs = {
            "fg_color": self.colors["panel"],
            "button_color": self.colors["border"],
            "button_hover_color": self.colors["border"],
            "text_color": self.colors["text"],
            "dropdown_fg_color": self.colors["panel"],
            "dropdown_hover_color": self.colors["sidebar"],
            "dropdown_text_color": self.colors["text"]
        }

        ctk.CTkLabel(
            prompt_card,
            text="待審檔案",
            font=self.fonts["small"],
            text_color=self.colors["muted"]
        ).pack(anchor="w", padx=18, pady=(0, 6))

        review_row = ctk.CTkFrame(prompt_card, fg_color="transparent")
        review_row.pack(fill="x", padx=18, pady=(0, 12))

        self.combo_review_var = ctk.StringVar(value="(無)")
        self.combo_review = ctk.CTkOptionMenu(
            review_row,
            values=["(無)"],
            height=34,
            font=self.fonts["body"],
            variable=self.combo_review_var,
            **menu_kwargs
        )
        self.combo_review.pack(side="left", fill="x", expand=True)

        gen_btn = ctk.CTkButton(
            review_row,
            text="生成提示詞",
            command=self.generate_prompt,
            height=34,
            width=120,
            font=self.fonts["small"],
            fg_color=self.colors["accent"],
            hover_color=self.colors["accent_hover"],
            text_color="#ffffff",
            corner_radius=14
        )
        gen_btn.pack(side="left", padx=(10, 0))

        self.prompt_display = ctk.CTkTextbox(
            prompt_card,
            height=160,
            font=self.fonts["body"],
            wrap="word",
            fg_color=self.colors["panel"],
            text_color=self.colors["text"],
            border_width=1,
            border_color=self.colors["border"]
        )
        self.prompt_display.pack(fill="both", expand=True, padx=18, pady=(0, 12))

        btn_frame = ctk.CTkFrame(prompt_card, fg_color="transparent")
        btn_frame.pack(fill="x", padx=18, pady=(0, 16))

        copy_btn = ctk.CTkButton(
            btn_frame,
            text="複製提示詞",
            command=self.copy_prompt,
            height=32,
            font=self.fonts["small"],
            fg_color=self.colors["panel"],
            hover_color=self.colors["border"],
            text_color=self.colors["text"],
            border_width=1,
            border_color=self.colors["border"],
            corner_radius=14
        )
        copy_btn.pack(side="left", fill="x", expand=True, padx=(0, 6))

        notebooklm_btn = ctk.CTkButton(
            btn_frame,
            text="開啟 NotebookLM",
            command=lambda: webbrowser.open(self.cfg.notebook_url),
            height=32,
            font=self.fonts["small"],
            fg_color=self.colors["panel"],
            hover_color=self.colors["border"],
            text_color=self.colors["text"],
            border_width=1,
            border_color=self.colors["border"],
            corner_radius=14
        )
        notebooklm_btn.pack(side="left", fill="x", expand=True, padx=6)

        clear_prompt_btn = ctk.CTkButton(
            btn_frame,
            text="清空",
            command=self.clear_prompt,
            height=32,
            font=self.fonts["small"],
            fg_color=self.colors["panel"],
            hover_color=self.colors["border"],
            text_color=self.colors["text"],
            border_width=1,
            border_color=self.colors["border"],
            corner_radius=14
        )
        clear_prompt_btn.pack(side="left", fill="x", expand=True, padx=(6, 0))

        header_actions = ctk.CTkFrame(reply_header, fg_color="transparent")
        header_actions.pack(side="right")

        clipboard_export_btn = ctk.CTkButton(
            header_actions,
            text="從剪貼簿輸出",
            command=self.export_word_from_clipboard,
            height=30,
            font=self.fonts["small"],
            fg_color=self.colors["panel"],
            hover_color=self.colors["border"],
            text_color=self.colors["text"],
            border_width=1,
            border_color=self.colors["border"],
            corner_radius=14
        )
        clipboard_export_btn.pack(side="left", padx=(0, 8))

        export_btn = ctk.CTkButton(
            header_actions,
            text="輸出 Word",
            command=self.export_word,
            height=30,
            font=self.fonts["small"],
            fg_color=self.colors["accent"],
            hover_color=self.colors["accent_hover"],
            text_color="#ffffff",
            corner_radius=14
        )
        export_btn.pack(side="left")

        self.reply_display = ctk.CTkTextbox(
            reply_card,
            font=self.fonts["body"],
            wrap="word",
            fg_color=self.colors["panel"],
            text_color=self.colors["text"],
            border_width=1,
            border_color=self.colors["border"]
        )
        self.reply_display.pack(fill="both", expand=True, padx=18, pady=(0, 12))

        reply_footer = ctk.CTkFrame(reply_card, fg_color="transparent")
        reply_footer.pack(fill="x", padx=18, pady=(0, 16))

        self.clipboard_auto_var = ctk.BooleanVar(value=False)
        clipboard_watch = ctk.CTkCheckBox(
            reply_footer,
            text="自動監聽剪貼簿並輸出",
            variable=self.clipboard_auto_var,
            command=self.toggle_clipboard_watch,
            font=self.fonts["small"],
            text_color=self.colors["muted"],
            fg_color=self.colors["accent"],
            hover_color=self.colors["accent_hover"]
        )
        clipboard_watch.pack(side="left")

        reply_actions = ctk.CTkFrame(reply_footer, fg_color="transparent")
        reply_actions.pack(side="right")

        open_output_btn = ctk.CTkButton(
            reply_actions,
            text="開啟 output",
            command=self.open_output,
            height=30,
            font=self.fonts["small"],
            fg_color=self.colors["panel"],
            hover_color=self.colors["border"],
            text_color=self.colors["text"],
            border_width=1,
            border_color=self.colors["border"],
            corner_radius=14
        )
        open_output_btn.pack(side="left", padx=(0, 8))

        clear_reply_btn = ctk.CTkButton(
            reply_actions,
            text="清空回覆",
            command=self.clear_reply,
            height=30,
            font=self.fonts["small"],
            fg_color=self.colors["panel"],
            hover_color=self.colors["border"],
            text_color=self.colors["text"],
            border_width=1,
            border_color=self.colors["border"],
            corner_radius=14
        )
        clear_reply_btn.pack(side="left")

    def _build_tagged_panel(self, parent):
        """右側：已標記檔案清單"""
        self.tagged_header_label = ctk.CTkLabel(
            parent,
            text="已標記檔案",
            font=self.fonts["small"],
            text_color=self.colors["muted"]
        )
        self.tagged_header_label.pack(anchor="w", padx=8, pady=(10, 6))

        files_container = ctk.CTkFrame(
            parent,
            fg_color=self.colors["panel"],
            corner_radius=14,
            border_width=1,
            border_color=self.colors["border"]
        )
        files_container.pack(fill="both", expand=True, padx=8, pady=(0, 12))

        self.tagged_list = ctk.CTkTextbox(
            files_container,
            font=self.fonts["small"],
            wrap="word",
            fg_color=self.colors["panel"],
            text_color=self.colors["text"],
            border_width=0
        )
        self.tagged_list.pack(fill="both", expand=True, padx=12, pady=12)

    def _show_file_tab(self, tab_name: str):
        if tab_name == "已標記":
            self.file_tab_inbox.pack_forget()
            self.file_tab_tagged.pack(fill="both", expand=True)
        else:
            self.file_tab_tagged.pack_forget()
            self.file_tab_inbox.pack(fill="both", expand=True)

    def _on_file_tab_change(self, value: str):
        self._show_file_tab(value)

    def _create_step_card(self, parent, title, description):
        card = ctk.CTkFrame(
            parent,
            fg_color=self.colors["panel"],
            corner_radius=18,
            border_width=1,
            border_color=self.colors["border"]
        )

        header = ctk.CTkFrame(card, fg_color="transparent")
        header.pack(fill="x", padx=18, pady=(16, 6))

        ctk.CTkLabel(
            header,
            text=title,
            font=self.fonts["section"],
            text_color=self.colors["text"],
            anchor="w"
        ).pack(side="left")

        ctk.CTkLabel(
            card,
            text=description,
            font=self.fonts["small"],
            text_color=self.colors["muted"],
            anchor="w"
        ).pack(anchor="w", padx=18, pady=(0, 10))

        return card, header

    def _create_file_item(self, parent, filename: str):
        """創建檔案項目（含三個標記按鈕）"""
        item = ctk.CTkFrame(
            parent,
            fg_color=self.colors["panel"],
            corner_radius=12,
            border_width=1,
            border_color=self.colors["border"],
            height=52
        )
        item.pack(fill="x", pady=5, padx=(0, 16))
        item.pack_propagate(False)

        # 檔名
        name_label = ctk.CTkLabel(
            item,
            text=filename,
            font=self.fonts["body"],
            text_color=self.colors["text"],
            anchor="w"
        )
        name_label.pack(side="left", padx=15, fill="x", expand=True)

        # 按鈕區
        btn_container = ctk.CTkFrame(item, fg_color="transparent")
        btn_container.pack(side="right", padx=10)

        colors = {
            "【標準】": ("#e6f0ff", "#d6e7ff", "#007aff"),
            "【範本】": ("#eaf7ef", "#dff2e7", "#34c759"),
            "【待審】": ("#fdecea", "#fbd9d6", "#ff3b30")
        }

        for tag in self.cfg.tags:
            fg, hover, text = colors[tag]
            btn = ctk.CTkButton(
                btn_container,
                text=tag,
                command=lambda t=tag, f=filename: self.tag_file(f, t),
                width=68,
                height=30,
                font=self.fonts["small"],
                fg_color=fg,
                hover_color=hover,
                text_color=text,
                corner_radius=12
            )
            btn.pack(side="left", padx=3)

    # ==================== 行為方法 ====================

    def current_project(self):
        return self.combo_project.get()

    def current_delivery(self):
        return self.combo_delivery.get()

    def on_selection_change(self):
        self.refresh_all()
        self._restart_watchdog()

    def open_input(self):
        path = self.fm.input_dir(self.current_project(), self.current_delivery())
        open_folder(path)

    def open_output(self):
        path = self.fm.output_dir(self.current_project(), self.current_delivery())
        open_folder(path)

    def show_help(self):
        help_text = (
            "1) 點「📂」開啟 input 資料夾，把檔案放進去\n"
            "2) 在左側為未標記檔案加上【標準/範本/待審】標籤\n"
            "3) 中間選擇待審檔案，生成提示詞並貼到 NotebookLM\n"
            "4) 把 AI 回覆貼回來，按「輸出 Word 報告」\n"
            "5) 或直接用「從剪貼簿輸出 / 自動監聽剪貼簿」"
        )
        messagebox.showinfo("使用說明", help_text)

    def tag_file(self, filename: str, tag: str):
        """標記檔案"""
        p, d = self.current_project(), self.current_delivery()
        success, result = self.fm.tag_file(p, d, filename, tag)
        
        if success:
            self.show_notification(f"✓ 已標記為 {tag}")
            self.refresh_all()
        else:
            messagebox.showerror("錯誤", f"標記失敗：{result}")

    def refresh_all(self):
        """刷新所有顯示"""
        p, d = self.current_project(), self.current_delivery()
        tagged, untagged = self.fm.list_input_files(p, d)
        self.refresh_untagged_files(untagged)
        self.refresh_tagged_files(tagged)
        self.refresh_review_combo(tagged)
        self._update_status(tagged, untagged)

    def refresh_untagged_files(self, untagged: Optional[List[str]] = None):
        """刷新未標記檔案列表"""
        # 清空容器
        for widget in self.untagged_container.winfo_children():
            widget.destroy()

        if untagged is None:
            p, d = self.current_project(), self.current_delivery()
            _, untagged = self.fm.list_input_files(p, d)

        if not untagged:
            empty_label = ctk.CTkLabel(
                self.untagged_container,
                text="沒有未標記的檔案\n\n把檔案放入 input 資料夾\n即可在這裡快速標記",
                font=self.fonts["small"],
                text_color=self.colors["muted"],
                justify="center"
            )
            empty_label.pack(pady=50)
        else:
            for filename in untagged:
                self._create_file_item(self.untagged_container, filename)

    def refresh_tagged_files(self, tagged: Optional[List[str]] = None):
        """刷新已標記檔案列表"""
        self.tagged_list.delete("1.0", "end")

        if tagged is None:
            p, d = self.current_project(), self.current_delivery()
            tagged, _ = self.fm.list_input_files(p, d)

        if not tagged:
            self.tagged_list.insert("end", "\n  尚無已標記檔案")
        else:
            std = [f for f in tagged if f.startswith("【標準】")]
            tpl = [f for f in tagged if f.startswith("【範本】")]
            rev = [f for f in tagged if f.startswith("【待審】")]

            if std:
                self.tagged_list.insert("end", "標準文件\n", "header")
                for f in std:
                    self.tagged_list.insert("end", f"  • {f}\n")
                self.tagged_list.insert("end", "\n")

            if tpl:
                self.tagged_list.insert("end", "範本文件\n", "header")
                for f in tpl:
                    self.tagged_list.insert("end", f"  • {f}\n")
                self.tagged_list.insert("end", "\n")

            if rev:
                self.tagged_list.insert("end", "待審文件\n", "header")
                for f in rev:
                    self.tagged_list.insert("end", f"  • {f}\n")

        # CTkTextbox forbids per-tag font to keep scaling consistent; use color only.
        self.tagged_list.tag_config("header", foreground=self.colors["accent"])

    def refresh_review_combo(self, tagged: Optional[List[str]] = None):
        """刷新待審檔案下拉選單"""
        if tagged is None:
            p, d = self.current_project(), self.current_delivery()
            tagged, _ = self.fm.list_input_files(p, d)

        review = [f for f in tagged if f.startswith("【待審】")]
        if review:
            self.combo_review.configure(values=review, state="normal")
            self.combo_review.set(review[0])
        else:
            self.combo_review.configure(values=["(無)"], state="disabled")
            self.combo_review.set("(無)")

    def _update_status(self, tagged: List[str], untagged: List[str]):
        if hasattr(self, "untagged_header_label"):
            self.untagged_header_label.configure(text=f"待標記檔案 ({len(untagged)})")
        if hasattr(self, "tagged_header_label"):
            self.tagged_header_label.configure(text=f"已標記檔案 ({len(tagged)})")

        p, d = self.current_project(), self.current_delivery()
        if hasattr(self, "status_left"):
            self.status_left.configure(
                text=f"專案：{p}  交付：{d}  |  待標記：{len(untagged)}  已標記：{len(tagged)}"
            )
        if hasattr(self, "status_right"):
            input_path = shorten_path(self.fm.input_dir(p, d), 52)
            output_path = shorten_path(self.fm.output_dir(p, d), 52)
            self.status_right.configure(text=f"input: {input_path}  |  output: {output_path}")

    def generate_prompt(self):
        """生成審查提示詞"""
        p, d = self.current_project(), self.current_delivery()
        tagged, untagged = self.fm.list_input_files(p, d)
        self.refresh_review_combo(tagged)
        
        std = [f for f in tagged if f.startswith("【標準】")]
        tpl = [f for f in tagged if f.startswith("【範本】")]
        tgt = self.combo_review_var.get()

        if tgt == "(無)":
            messagebox.showwarning("提醒", "沒有待審檔案")
            return

        prompt = f"""請以【標準】與【範本】作為依據，逐條審查【待審】文件：{tgt}

【標準】
""" + "\n".join(f"- {x}" for x in std or ["(無)"]) + """

【範本】
""" + "\n".join(f"- {x}" for x in tpl or ["(無)"]) + """

請輸出：
1) 不符合之處
2) 風險
3) 具體修改建議
4) 需人工確認事項
"""

        self.prompt_display.delete("1.0", "end")
        self.prompt_display.insert("1.0", prompt)
        self.show_notification("✓ 提示詞已生成")

    def copy_prompt(self):
        """複製提示詞"""
        txt = self.prompt_display.get("1.0", "end").strip()
        if not txt:
            messagebox.showwarning("提醒", "提示詞是空的")
            return
        self.clipboard_clear()
        self.clipboard_append(txt)
        self.show_notification("✓ 已複製到剪貼簿")

    def clear_prompt(self):
        self.prompt_display.delete("1.0", "end")
        self.show_notification("✓ 提示詞已清空")

    def clear_reply(self):
        self.reply_display.delete("1.0", "end")
        self.show_notification("✓ 回覆內容已清空")

    def _get_clipboard_text(self) -> str:
        try:
            text = self.clipboard_get()
        except TclError:
            return ""
        if not isinstance(text, str):
            return ""
        return text.strip()

    def _set_reply_text(self, text: str):
        self.reply_display.delete("1.0", "end")
        self.reply_display.insert("1.0", text)

    def export_word(self):
        """輸出 Word 報告"""
        content = self.reply_display.get("1.0", "end").strip()
        self._export_content(content, open_dir=True, show_error_dialog=True)

    def export_word_from_clipboard(self):
        """從剪貼簿輸出 Word 報告"""
        content = self._get_clipboard_text()
        if not content:
            messagebox.showwarning("提醒", "剪貼簿沒有文字內容")
            return
        self._set_reply_text(content)
        self._export_content(content, open_dir=True, show_error_dialog=True)

    def _export_content(self, content: str, open_dir: bool, show_error_dialog: bool):
        tgt = self.combo_review_var.get()

        if not content:
            if show_error_dialog:
                messagebox.showwarning("提醒", "回覆內容是空的")
            else:
                self.show_notification("⚠ 回覆內容是空的")
            return False

        if tgt == "(無)":
            if show_error_dialog:
                messagebox.showwarning("提醒", "沒有選擇待審檔案")
            else:
                self.show_notification("⚠ 請先選擇待審檔案")
            return False

        try:
            path = self.exporter.export(
                self.fm.output_dir(self.current_project(), self.current_delivery()),
                tgt,
                content
            )
            self.show_notification(f"✓ Word 已輸出：{os.path.basename(path)}")
            if open_dir:
                open_folder(os.path.dirname(path))
            return True
        except Exception as e:
            self.logger.error("Word 輸出失敗：%s", e)
            if show_error_dialog:
                messagebox.showerror("錯誤", f"Word 輸出失敗：{e}")
            else:
                self.show_notification("⚠ Word 輸出失敗")
            return False

    def show_notification(self, message: str):
        """顯示通知"""
        self.notification.configure(text=f"  {message}  ", height=36)
        self.after(3000, lambda: self.notification.configure(text="", height=0))

    def toggle_clipboard_watch(self):
        if self.clipboard_auto_var.get():
            self._last_clipboard = None
            self.show_notification("✓ 已啟用剪貼簿監聽")
            self._schedule_clipboard_poll(immediate=True)
        else:
            if self._clipboard_job:
                self.after_cancel(self._clipboard_job)
                self._clipboard_job = None
            self.show_notification("✓ 已停止剪貼簿監聽")

    def _schedule_clipboard_poll(self, immediate: bool = False):
        if self._clipboard_job:
            self.after_cancel(self._clipboard_job)
        delay = 10 if immediate else self._clipboard_poll_ms
        self._clipboard_job = self.after(delay, self._poll_clipboard)

    def _poll_clipboard(self):
        self._clipboard_job = None
        if not self.clipboard_auto_var.get():
            return
        content = self._get_clipboard_text()
        if content and content != self._last_clipboard:
            self._last_clipboard = content
            self._set_reply_text(content)
            self._export_content(content, open_dir=False, show_error_dialog=False)
        self._schedule_clipboard_poll()

    def schedule_refresh(self, delay_ms: int = 300):
        if self._refresh_job:
            self.after_cancel(self._refresh_job)
        self._refresh_job = self.after(delay_ms, self._run_refresh)

    def _run_refresh(self):
        self._refresh_job = None
        self.refresh_all()

    def _start_watchdog(self):
        """啟動檔案監控"""
        if self.observer is None:
            self.observer = Observer()
            self.observer.start()
        self._restart_watchdog()

    def _restart_watchdog(self):
        """重新啟動檔案監控"""
        if self.observer:
            self.observer.unschedule_all()
            watch_path = self.fm.input_dir(self.current_project(), self.current_delivery())
            handler = AutoTagHandler(self)
            self.observer.schedule(handler, watch_path, recursive=False)

    def destroy(self):
        """清理資源"""
        if self._clipboard_job:
            self.after_cancel(self._clipboard_job)
            self._clipboard_job = None
        if self.observer:
            self.observer.stop()
            self.observer.join()
        super().destroy()

# ==================== main ====================

if __name__ == "__main__":
    app = NotebookLMSingleFolderApp()
    app.mainloop()
