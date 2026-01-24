#!/usr/bin/env python3
"""
PDF Crop Tool v2
A GUI application for managing PDF/PNG sources and building worksheet projects.

Two modes:
- Source Mode: Import PDFs/PNG folders, set crops, tag pages
- Project Mode: Build worksheets from sources, filter by tags, rearrange, export
"""

import base64
import io
import json
import os
import random
import shutil
import subprocess
import sys
import tempfile
import threading
import tkinter as tk
import urllib.request
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog
from typing import Optional

OPENROUTER_API_KEY = None

VISION_MODELS = [
    ("Gemini 3 Flash Preview", "google/gemini-3-flash-preview"),
    ("Gemini 3 Pro Preview", "google/gemini-3-pro-preview"),
    ("Claude Sonnet 4.5", "anthropic/claude-sonnet-4.5"),
    ("Claude Opus 4.5", "anthropic/claude-opus-4.5"),
    ("GPT-5.1", "openai/gpt-5.1"),
    ("GPT-4.5 Preview", "openai/gpt-4.5-preview"),
]

import customtkinter as ctk
import fitz  # PyMuPDF
from PIL import Image, ImageDraw, ImageEnhance, ImageTk

# Config files
CONFIG_DIR = Path.home() / ".pdf_crop_tool"
CONFIG_DIR.mkdir(exist_ok=True)

CONFIG_FILE = CONFIG_DIR / "config.json"
SOURCES_INDEX_FILE = CONFIG_DIR / "sources_index.json"
PROJECTS_INDEX_FILE = CONFIG_DIR / "projects_index.json"


def load_json(path: Path, default=None):
    if default is None:
        default = {}
    if path.exists():
        try:
            return json.loads(path.read_text())
        except (json.JSONDecodeError, IOError):
            return default
    return default


def save_json(path: Path, data):
    path.write_text(json.dumps(data, indent=2))


def copy_image_to_clipboard(img: Image.Image):
    temp_path = os.path.join(tempfile.gettempdir(), "pdf_crop_clipboard.png")
    img.save(temp_path, format="PNG")
    
    if sys.platform == "darwin":
        subprocess.run([
            "osascript", "-e",
            f'set the clipboard to (read (POSIX file "{temp_path}") as «class PNGf»)'
        ], check=True)
    elif sys.platform == "win32":
        temp_bmp = os.path.join(tempfile.gettempdir(), "pdf_crop_clipboard.bmp")
        img.convert("RGB").save(temp_bmp, format="BMP")
        
        ps_script = f'''
Add-Type -AssemblyName System.Windows.Forms
$img = [System.Drawing.Image]::FromFile("{temp_bmp.replace(chr(92), '/')}")
[System.Windows.Forms.Clipboard]::SetImage($img)
$img.Dispose()
'''
        subprocess.run(["powershell", "-Command", ps_script], check=True, capture_output=True)
    else:
        raise Exception("Clipboard not supported on this platform")


class AppConfig:
    """Global app configuration."""
    
    def __init__(self):
        self.data = load_json(CONFIG_FILE, {
            "sources_folder": None,
            "projects_folder": None,
            "zoom": 0.5,
            "theme": "dark",
            "recent_sources": [],
            "recent_projects": [],
            "last_export_folder": None,
            "presets": {}
        })
    
    def save(self):
        save_json(CONFIG_FILE, self.data)
    
    def add_recent_source(self, path: Path):
        path_str = str(path)
        recents = self.data.get("recent_sources", [])
        if path_str in recents:
            recents.remove(path_str)
        recents.insert(0, path_str)
        self.data["recent_sources"] = recents[:10]
        self.save()
    
    def add_recent_project(self, path: Path):
        path_str = str(path)
        recents = self.data.get("recent_projects", [])
        if path_str in recents:
            recents.remove(path_str)
        recents.insert(0, path_str)
        self.data["recent_projects"] = recents[:10]
        self.save()
    
    def get_recent_sources(self) -> list[Path]:
        return [Path(p) for p in self.data.get("recent_sources", []) if Path(p).exists()]
    
    def get_recent_projects(self) -> list[Path]:
        return [Path(p) for p in self.data.get("recent_projects", []) if Path(p).exists()]
    
    @property
    def last_export_folder(self) -> Optional[Path]:
        if self.data.get("last_export_folder"):
            p = Path(self.data["last_export_folder"])
            if p.exists():
                return p
        return None
    
    @last_export_folder.setter
    def last_export_folder(self, path: Path):
        self.data["last_export_folder"] = str(path)
        self.save()
    
    def get_presets(self) -> dict:
        return self.data.get("presets", {})
    
    def save_preset(self, name: str, crop: dict):
        if "presets" not in self.data:
            self.data["presets"] = {}
        self.data["presets"][name] = crop
        self.save()
    
    def delete_preset(self, name: str):
        if "presets" in self.data and name in self.data["presets"]:
            del self.data["presets"][name]
            self.save()
    
    def add_history(self, entry: dict):
        if "history" not in self.data:
            self.data["history"] = []
        entry["timestamp"] = datetime.now().isoformat()
        self.data["history"].insert(0, entry)
        self.data["history"] = self.data["history"][:500]
        self.save()
    
    def get_history(self) -> list[dict]:
        return self.data.get("history", [])
    
    @property
    def api_key(self) -> Optional[str]:
        return self.data.get("openrouter_api_key")
    
    @api_key.setter
    def api_key(self, key: str):
        self.data["openrouter_api_key"] = key
        self.save()
    
    @property
    def sources_folder(self) -> Optional[Path]:
        if self.data.get("sources_folder"):
            return Path(self.data["sources_folder"])
        return None
    
    @sources_folder.setter
    def sources_folder(self, path: Path):
        self.data["sources_folder"] = str(path)
        self.save()
    
    @property
    def projects_folder(self) -> Optional[Path]:
        if self.data.get("projects_folder"):
            return Path(self.data["projects_folder"])
        return None
    
    @projects_folder.setter
    def projects_folder(self, path: Path):
        self.data["projects_folder"] = str(path)
        self.save()


class Source:
    """Represents a source (PDF or folder of PNGs) with metadata."""
    
    def __init__(self, source_path: Path):
        self.path = source_path
        self.meta_file = source_path / "source_meta.json"
        self.meta = self._load_meta()
    
    def _load_meta(self) -> dict:
        return load_json(self.meta_file, {
            "name": self.path.name,
            "type": "pdf",
            "original_file": None,
            "page_range": {"start": 1, "end": None},
            "default_crop": {"left": 0, "right": 0, "top": 0, "bottom": 0},
            "page_tags": {},
            "page_crops": {},
            "created": None,
            "zoom": 0.5,
            "last_page": 1,
        })
    
    def save_meta(self):
        save_json(self.meta_file, self.meta)
    
    @property
    def name(self) -> str:
        return self.meta.get("name", self.path.name)
    
    @property
    def source_type(self) -> str:
        return self.meta.get("type", "pdf")
    
    @property
    def pdf_path(self) -> Optional[Path]:
        if self.source_type == "pdf":
            for f in self.path.iterdir():
                if f.suffix.lower() == ".pdf":
                    return f
        return None
    
    def get_page_count(self) -> int:
        if self.source_type == "pdf":
            pdf_path = self.pdf_path
            if pdf_path and pdf_path.exists():
                doc = fitz.open(str(pdf_path))
                count = len(doc)
                doc.close()
                return count
        else:
            # PNG folder
            pngs = list(self.path.glob("*.png")) + list(self.path.glob("*.PNG"))
            return len(pngs)
        return 0
    
    def get_page_range(self) -> tuple[int, int]:
        pr = self.meta.get("page_range", {})
        start = pr.get("start", 1)
        end = pr.get("end") or self.get_page_count()
        return start, end
    
    def get_all_tags(self) -> set[str]:
        tags = set()
        for page_tags in self.meta.get("page_tags", {}).values():
            tags.update(page_tags)
        return tags
    
    def get_page_tags(self, page_num: int) -> list[str]:
        return self.meta.get("page_tags", {}).get(str(page_num), [])
    
    def set_page_tags(self, page_num: int, tags: list[str]):
        if "page_tags" not in self.meta:
            self.meta["page_tags"] = {}
        self.meta["page_tags"][str(page_num)] = tags
        self.save_meta()
    
    def add_page_tag(self, page_num: int, tag: str):
        tags = self.get_page_tags(page_num)
        if tag not in tags:
            tags.append(tag)
            self.set_page_tags(page_num, tags)
    
    def remove_page_tag(self, page_num: int, tag: str):
        tags = self.get_page_tags(page_num)
        if tag in tags:
            tags.remove(tag)
            self.set_page_tags(page_num, tags)
    
    def get_default_crop(self) -> dict:
        return self.meta.get("default_crop", {"left": 0, "right": 0, "top": 0, "bottom": 0})
    
    def set_default_crop(self, crop: dict):
        self.meta["default_crop"] = crop
        self.save_meta()
    
    def get_page_crop(self, page_num: int) -> dict:
        override = self.meta.get("page_crops", {}).get(str(page_num))
        if override:
            return override
        return self.get_default_crop()
    
    def set_page_crop(self, page_num: int, crop: dict):
        if "page_crops" not in self.meta:
            self.meta["page_crops"] = {}
        self.meta["page_crops"][str(page_num)] = crop
        self.save_meta()
    
    def clear_page_crop(self, page_num: int):
        if "page_crops" in self.meta and str(page_num) in self.meta["page_crops"]:
            del self.meta["page_crops"][str(page_num)]
            self.save_meta()
    
    def has_page_crop_override(self, page_num: int) -> bool:
        return str(page_num) in self.meta.get("page_crops", {})
    
    def get_tag_definitions(self) -> list[dict]:
        return self.meta.get("tag_definitions", [])
    
    def set_tag_definitions(self, definitions: list[dict]):
        self.meta["tag_definitions"] = definitions
        self.save_meta()
    
    def add_tag_definition(self, name: str, description: str):
        defs = self.get_tag_definitions()
        for d in defs:
            if d["name"] == name:
                d["description"] = description
                self.set_tag_definitions(defs)
                return
        defs.append({"name": name, "description": description})
        self.set_tag_definitions(defs)
    
    def remove_tag_definition(self, name: str):
        defs = [d for d in self.get_tag_definitions() if d["name"] != name]
        self.set_tag_definitions(defs)


class Project:
    """Represents a worksheet project."""
    
    def __init__(self, project_path: Path):
        self.path = project_path
        self.meta_file = project_path / "project_meta.json"
        self.meta = self._load_meta()
    
    def _load_meta(self) -> dict:
        return load_json(self.meta_file, {
            "name": self.path.name,
            "created": None,
            "pages": [],  # [{"source": "path", "page": 1, "type": "source"}, {"path": "custom.png", "type": "custom"}]
        })
    
    def save_meta(self):
        save_json(self.meta_file, self.meta)
    
    @property
    def name(self) -> str:
        return self.meta.get("name", self.path.name)
    
    @property
    def pages(self) -> list[dict]:
        return self.meta.get("pages", [])
    
    def add_page(self, page_info: dict):
        self.meta["pages"].append(page_info)
        self.save_meta()
    
    def add_pages(self, pages: list[dict]):
        self.meta["pages"].extend(pages)
        self.save_meta()
    
    def remove_page(self, index: int):
        if 0 <= index < len(self.meta["pages"]):
            del self.meta["pages"][index]
            self.save_meta()
    
    def move_page(self, from_idx: int, to_idx: int):
        pages = self.meta["pages"]
        if 0 <= from_idx < len(pages) and 0 <= to_idx < len(pages):
            page = pages.pop(from_idx)
            pages.insert(to_idx, page)
            self.save_meta()
    
    def clear_pages(self):
        self.meta["pages"] = []
        self.save_meta()


class WelcomeScreen(ctk.CTkFrame):
    """Initial screen shown when app starts - choose mode or first-time setup."""
    
    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self.config = app.config
        
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        
        center_frame = ctk.CTkFrame(self, fg_color="transparent")
        center_frame.grid(row=0, column=0)
        
        title = ctk.CTkLabel(
            center_frame,
            text="PDF Crop Tool",
            font=ctk.CTkFont(size=32, weight="bold")
        )
        title.pack(pady=(0, 10))
        
        subtitle = ctk.CTkLabel(
            center_frame,
            text="Manage sources and build worksheet projects",
            font=ctk.CTkFont(size=14),
            text_color="#888888"
        )
        subtitle.pack(pady=(0, 40))
        
        # Check if first-time setup needed
        if not self.config.sources_folder:
            self._show_first_time_setup(center_frame)
        else:
            self._show_mode_selection(center_frame)
    
    def _show_first_time_setup(self, parent):
        setup_label = ctk.CTkLabel(
            parent,
            text="First-time setup: Choose your workspace folder",
            font=ctk.CTkFont(size=16)
        )
        setup_label.pack(pady=(0, 10))
        
        ctk.CTkLabel(
            parent,
            text="This will create 'sources' and 'projects' subfolders",
            text_color="#888888"
        ).pack(pady=(0, 20))
        
        workspace_frame = ctk.CTkFrame(parent, fg_color="transparent")
        workspace_frame.pack(pady=10)
        
        ctk.CTkLabel(workspace_frame, text="Workspace:").pack(side="left", padx=(0, 10))
        
        self.workspace_path_label = ctk.CTkLabel(
            workspace_frame,
            text="Not selected",
            text_color="#888888",
            width=300
        )
        self.workspace_path_label.pack(side="left", padx=(0, 10))
        
        ctk.CTkButton(
            workspace_frame,
            text="Browse",
            width=80,
            command=self._select_workspace_folder
        ).pack(side="left")
        
        self.continue_btn = ctk.CTkButton(
            parent,
            text="Continue",
            width=200,
            height=40,
            state="disabled",
            command=self._complete_setup
        )
        self.continue_btn.pack(pady=30)
    
    def _select_workspace_folder(self):
        folder = filedialog.askdirectory(title="Select Workspace Folder")
        if folder:
            self.selected_workspace = Path(folder)
            self.workspace_path_label.configure(text=str(self.selected_workspace))
            self.continue_btn.configure(state="normal")
    
    def _complete_setup(self):
        sources_folder = self.selected_workspace / "sources"
        projects_folder = self.selected_workspace / "projects"
        
        sources_folder.mkdir(parents=True, exist_ok=True)
        projects_folder.mkdir(parents=True, exist_ok=True)
        
        self.config.sources_folder = sources_folder
        self.config.projects_folder = projects_folder
        self.app.show_welcome()
    
    def _show_mode_selection(self, parent):
        btn_frame = ctk.CTkFrame(parent, fg_color="transparent")
        btn_frame.pack()
        
        source_btn = ctk.CTkButton(
            btn_frame,
            text="Work on Sources",
            width=200,
            height=60,
            font=ctk.CTkFont(size=16),
            command=lambda: self.app.show_source_browser()
        )
        source_btn.pack(side="left", padx=20)
        
        project_btn = ctk.CTkButton(
            btn_frame,
            text="Work on Projects",
            width=200,
            height=60,
            font=ctk.CTkFont(size=16),
            command=lambda: self.app.show_project_browser()
        )
        project_btn.pack(side="left", padx=20)
        
        desc_frame = ctk.CTkFrame(parent, fg_color="transparent")
        desc_frame.pack(pady=(10, 0))
        
        ctk.CTkLabel(
            desc_frame,
            text="Import PDFs, set crops, tag pages",
            text_color="#888888",
            width=200
        ).pack(side="left", padx=20)
        
        ctk.CTkLabel(
            desc_frame,
            text="Build worksheets from sources",
            text_color="#888888",
            width=200
        ).pack(side="left", padx=20)
        
        self._show_recent_items(parent)
        
        settings_frame = ctk.CTkFrame(parent, fg_color="transparent")
        settings_frame.pack(pady=(30, 0))
        
        ctk.CTkButton(
            settings_frame,
            text="Change Sources Folder",
            width=160,
            fg_color="transparent",
            border_width=1,
            command=self._change_sources_folder
        ).pack(side="left", padx=5)
    
    def _show_recent_items(self, parent):
        recent_sources = self.config.get_recent_sources()
        recent_projects = self.config.get_recent_projects()
        
        if not recent_sources and not recent_projects:
            return
        
        recent_frame = ctk.CTkFrame(parent, fg_color="transparent")
        recent_frame.pack(pady=(30, 0), fill="x", padx=50)
        
        if recent_sources:
            src_frame = ctk.CTkFrame(recent_frame)
            src_frame.pack(side="left", padx=10, fill="both", expand=True)
            
            ctk.CTkLabel(
                src_frame,
                text="Recent Sources",
                font=ctk.CTkFont(size=14, weight="bold")
            ).pack(pady=(10, 5), padx=10, anchor="w")
            
            for source_path in recent_sources[:5]:
                if source_path.exists():
                    source = Source(source_path)
                    btn = ctk.CTkButton(
                        src_frame,
                        text=f"📄 {source.name}",
                        anchor="w",
                        fg_color="transparent",
                        text_color=("#000", "#fff"),
                        hover_color=("#e0e0e0", "#3a3a3a"),
                        command=lambda s=source: self.app.show_source_editor(s)
                    )
                    btn.pack(fill="x", padx=10, pady=2)
        
        if recent_projects:
            proj_frame = ctk.CTkFrame(recent_frame)
            proj_frame.pack(side="left", padx=10, fill="both", expand=True)
            
            ctk.CTkLabel(
                proj_frame,
                text="Recent Projects",
                font=ctk.CTkFont(size=14, weight="bold")
            ).pack(pady=(10, 5), padx=10, anchor="w")
            
            for project_path in recent_projects[:5]:
                if project_path.exists():
                    project = Project(project_path)
                    btn = ctk.CTkButton(
                        proj_frame,
                        text=f"📋 {project.name}",
                        anchor="w",
                        fg_color="transparent",
                        text_color=("#000", "#fff"),
                        hover_color=("#e0e0e0", "#3a3a3a"),
                        command=lambda p=project: self.app.show_project_editor(p)
                    )
                    btn.pack(fill="x", padx=10, pady=2)
    
    def _change_sources_folder(self):
        folder = filedialog.askdirectory(
            title="Select Sources Folder",
            initialdir=str(self.config.sources_folder) if self.config.sources_folder else None
        )
        if folder:
            self.config.sources_folder = Path(folder)
            messagebox.showinfo("Updated", f"Sources folder set to:\n{folder}")


class SourceBrowser(ctk.CTkFrame):
    """Browse and manage sources."""
    
    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self.config = app.config
        self.search_var = tk.StringVar()
        self.all_sources: list[Source] = []
        self.expanded_folders: set[str] = set()
        
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)
        
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=20, pady=(20, 10))
        header.grid_columnconfigure(1, weight=1)
        
        ctk.CTkButton(
            header,
            text="< Back",
            width=80,
            command=self.app.show_welcome
        ).grid(row=0, column=0, padx=(0, 20))
        
        ctk.CTkLabel(
            header,
            text="Sources",
            font=ctk.CTkFont(size=24, weight="bold")
        ).grid(row=0, column=1, sticky="w")
        
        ctk.CTkButton(
            header,
            text="+ Add Source",
            width=120,
            fg_color="#28a745",
            hover_color="#218838",
            command=self._add_source
        ).grid(row=0, column=2)
        
        search_frame = ctk.CTkFrame(self, fg_color="transparent")
        search_frame.grid(row=1, column=0, sticky="ew", padx=20, pady=(0, 10))
        
        ctk.CTkLabel(search_frame, text="Search:").pack(side="left", padx=(0, 10))
        
        self.search_entry = ctk.CTkEntry(
            search_frame,
            textvariable=self.search_var,
            placeholder_text="Filter sources by name or tag...",
            width=400
        )
        self.search_entry.pack(side="left", fill="x", expand=True)
        self.search_var.trace_add("write", lambda *args: self._filter_sources())
        
        self.sources_frame = ctk.CTkScrollableFrame(self)
        self.sources_frame.grid(row=2, column=0, sticky="nsew", padx=20, pady=(0, 20))
        self.sources_frame.grid_columnconfigure(0, weight=1)
        
        self._refresh_sources()
    
    def _refresh_sources(self):
        self.all_sources.clear()
        
        sources_folder = self.config.sources_folder
        if sources_folder and sources_folder.exists():
            self._find_sources_recursive(sources_folder)
        
        self._render_sources()
    
    def _find_sources_recursive(self, folder: Path):
        try:
            for item in folder.iterdir():
                if item.is_dir():
                    if (item / "source_meta.json").exists():
                        self.all_sources.append(Source(item))
                    else:
                        self._find_sources_recursive(item)
        except PermissionError:
            pass
    
    def _render_sources(self):
        for widget in self.sources_frame.winfo_children():
            widget.destroy()
        
        if not self.config.sources_folder:
            ctk.CTkLabel(
                self.sources_frame,
                text="No sources folder configured",
                text_color="#888888"
            ).pack(pady=50)
            return
        
        if not self.all_sources:
            ctk.CTkLabel(
                self.sources_frame,
                text="No sources yet. Click '+ Add Source' to import a PDF or PNG folder.",
                text_color="#888888"
            ).pack(pady=50)
            return
        
        query = self.search_var.get().lower().strip()
        
        if query:
            filtered = []
            for source in self.all_sources:
                name_match = query in source.name.lower()
                tag_match = any(query in tag.lower() for tag in source.get_all_tags())
                path_match = query in str(source.path).lower()
                if name_match or tag_match or path_match:
                    filtered.append(source)
            
            if not filtered:
                ctk.CTkLabel(
                    self.sources_frame,
                    text=f"No sources match '{query}'",
                    text_color="#888888"
                ).pack(pady=50)
                return
            
            for source in sorted(filtered, key=lambda s: s.name.lower()):
                self._create_source_card(source, indent=0)
        else:
            self._render_folder_tree(self.config.sources_folder, indent=0)
    
    def _render_folder_tree(self, folder: Path, indent: int):
        try:
            items = sorted(folder.iterdir(), key=lambda x: (not x.is_dir(), x.name.lower()))
        except PermissionError:
            return
        
        for item in items:
            if not item.is_dir() or item.name.startswith('.'):
                continue
            
            if (item / "source_meta.json").exists():
                source = Source(item)
                self._create_source_card(source, indent)
            else:
                has_sources = self._folder_has_sources(item)
                if has_sources:
                    self._create_folder_row(item, indent)
                    
                    if str(item) in self.expanded_folders:
                        self._render_folder_tree(item, indent + 1)
    
    def _folder_has_sources(self, folder: Path) -> bool:
        try:
            for item in folder.iterdir():
                if item.is_dir():
                    if (item / "source_meta.json").exists():
                        return True
                    if self._folder_has_sources(item):
                        return True
        except PermissionError:
            pass
        return False
    
    def _count_sources_in_folder(self, folder: Path) -> int:
        count = 0
        try:
            for item in folder.iterdir():
                if item.is_dir():
                    if (item / "source_meta.json").exists():
                        count += 1
                    else:
                        count += self._count_sources_in_folder(item)
        except PermissionError:
            pass
        return count
    
    def _create_folder_row(self, folder: Path, indent: int):
        is_expanded = str(folder) in self.expanded_folders
        arrow = "▼" if is_expanded else "▶"
        count = self._count_sources_in_folder(folder)
        
        row = ctk.CTkFrame(self.sources_frame, fg_color="transparent")
        row.pack(fill="x", pady=2)
        
        btn = ctk.CTkButton(
            row,
            text=f"{arrow} 📁 {folder.name} ({count})",
            anchor="w",
            fg_color="transparent",
            text_color=("#333", "#ccc"),
            hover_color=("#e0e0e0", "#3a3a3a"),
            font=ctk.CTkFont(size=14),
            command=lambda f=folder: self._toggle_folder(f)
        )
        btn.pack(side="left", padx=(indent * 20, 0), fill="x", expand=True)
    
    def _toggle_folder(self, folder: Path):
        folder_str = str(folder)
        if folder_str in self.expanded_folders:
            self.expanded_folders.remove(folder_str)
        else:
            self.expanded_folders.add(folder_str)
        self._render_sources()
    
    def _filter_sources(self):
        self._render_sources()
    
    def _create_source_card(self, source: Source, indent: int = 0):
        card = ctk.CTkFrame(self.sources_frame)
        card.pack(fill="x", pady=3, padx=(indent * 20, 0))
        card.grid_columnconfigure(1, weight=1)
        
        icon = "📄" if source.source_type == "pdf" else "🖼️"
        ctk.CTkLabel(card, text=icon, font=ctk.CTkFont(size=20)).grid(
            row=0, column=0, rowspan=2, padx=10, pady=8
        )
        
        ctk.CTkLabel(
            card,
            text=source.name,
            font=ctk.CTkFont(size=14, weight="bold"),
            anchor="w"
        ).grid(row=0, column=1, sticky="w", pady=(8, 0))
        
        page_count = source.get_page_count()
        start, end = source.get_page_range()
        tags = source.get_all_tags()
        tag_str = f" | {', '.join(sorted(tags)[:3])}" if tags else ""
        if len(tags) > 3:
            tag_str += "..."
        
        ctk.CTkLabel(
            card,
            text=f"{page_count} pages ({start}-{end or page_count}){tag_str}",
            text_color="#888888",
            anchor="w",
            font=ctk.CTkFont(size=12)
        ).grid(row=1, column=1, sticky="w", pady=(0, 8))
        
        btn_frame = ctk.CTkFrame(card, fg_color="transparent")
        btn_frame.grid(row=0, column=2, rowspan=2, padx=10)
        
        ctk.CTkButton(
            btn_frame,
            text="Open",
            width=70,
            command=lambda s=source: self.app.show_source_editor(s)
        ).pack(side="left", padx=2)
        
        ctk.CTkButton(
            btn_frame,
            text="⋮",
            width=30,
            fg_color="transparent",
            border_width=1,
            command=lambda s=source: self._show_source_menu(s)
        ).pack(side="left", padx=2)
    
    def _show_source_menu(self, source: Source):
        SourceContextMenu(self, source, self._refresh_sources, self.app)
    
    def _add_source(self):
        AddSourceDialog(self, self.app, self._refresh_sources)


class SourceContextMenu(ctk.CTkToplevel):
    def __init__(self, parent, source: Source, callback, app):
        super().__init__(parent)
        self.source = source
        self.callback = callback
        self.app = app
        self.parent = parent
        
        self.title("")
        self.geometry("160x140")
        self.resizable(False, False)
        self.overrideredirect(True)
        
        x = parent.winfo_pointerx()
        y = parent.winfo_pointery()
        self.geometry(f"160x140+{x}+{y}")
        
        self.transient(parent)
        
        frame = ctk.CTkFrame(self, border_width=1, border_color="#555")
        frame.pack(fill="both", expand=True)
        
        ctk.CTkButton(
            frame, text="Rename", anchor="w", fg_color="transparent",
            text_color=("#000", "#fff"), hover_color=("#e0e0e0", "#3a3a3a"),
            command=self._rename
        ).pack(fill="x", padx=5, pady=(5, 2))
        
        ctk.CTkButton(
            frame, text="Duplicate", anchor="w", fg_color="transparent",
            text_color=("#000", "#fff"), hover_color=("#e0e0e0", "#3a3a3a"),
            command=self._duplicate
        ).pack(fill="x", padx=5, pady=2)
        
        ctk.CTkButton(
            frame, text="Delete", anchor="w", fg_color="transparent",
            text_color="#ff6666", hover_color=("#ffcccc", "#4a2020"),
            command=self._delete
        ).pack(fill="x", padx=5, pady=2)
        
        ctk.CTkButton(
            frame, text="Cancel", anchor="w", fg_color="transparent",
            text_color="#888", hover_color=("#e0e0e0", "#3a3a3a"),
            command=self.destroy
        ).pack(fill="x", padx=5, pady=(2, 5))
        
        self.bind("<Escape>", lambda e: self.destroy())
        self.bind("<FocusOut>", lambda e: self.destroy())
        
        self.focus_set()
    
    def _rename(self):
        self.destroy()
        new_name = simpledialog.askstring("Rename Source", "New name:", initialvalue=self.source.name)
        if new_name and new_name.strip():
            new_name = new_name.strip()
            self.source.meta["name"] = new_name
            self.source.save_meta()
            self.callback()
    
    def _duplicate(self):
        self.destroy()
        new_name = simpledialog.askstring("Duplicate Source", "Name for copy:", 
                                          initialvalue=f"{self.source.name} (copy)")
        if new_name and new_name.strip():
            new_name = new_name.strip()
            new_path = self.source.path.parent / new_name
            if new_path.exists():
                messagebox.showerror("Error", f"'{new_name}' already exists.")
                return
            try:
                shutil.copytree(self.source.path, new_path)
                new_source = Source(new_path)
                new_source.meta["name"] = new_name
                new_source.meta["created"] = datetime.now().isoformat()
                new_source.save_meta()
                self.callback()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to duplicate: {e}")
    
    def _delete(self):
        self.destroy()
        if messagebox.askyesno("Delete Source", 
                               f"Delete '{self.source.name}' and all its data?\n\nThis cannot be undone."):
            try:
                shutil.rmtree(self.source.path)
                self.callback()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to delete: {e}")


class AddSourceDialog(ctk.CTkToplevel):
    """Dialog to add a new source."""
    
    def __init__(self, parent, app, callback):
        super().__init__(parent)
        self.app = app
        self.config = app.config
        self.callback = callback
        
        self.title("Add Source")
        self.geometry("600x500")
        self.transient(parent)
        self.grab_set()
        
        self.grid_columnconfigure(0, weight=1)
        
        # Step 1: Select source file/folder
        step1 = ctk.CTkFrame(self)
        step1.grid(row=0, column=0, sticky="ew", padx=20, pady=20)
        step1.grid_columnconfigure(1, weight=1)
        
        ctk.CTkLabel(
            step1,
            text="1. Select Source",
            font=ctk.CTkFont(size=16, weight="bold")
        ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 10))
        
        self.source_path_label = ctk.CTkLabel(step1, text="No file selected", text_color="#888888")
        self.source_path_label.grid(row=1, column=0, columnspan=2, sticky="w", pady=5)
        
        btn_frame = ctk.CTkFrame(step1, fg_color="transparent")
        btn_frame.grid(row=2, column=0, columnspan=3, sticky="w")
        
        ctk.CTkButton(btn_frame, text="Select PDF", width=100, command=self._select_pdf).pack(side="left", padx=(0, 10))
        ctk.CTkButton(btn_frame, text="Select PNG Folder", width=140, command=self._select_png_folder).pack(side="left")
        
        # Step 2: Choose destination
        step2 = ctk.CTkFrame(self)
        step2.grid(row=1, column=0, sticky="ew", padx=20, pady=10)
        step2.grid_columnconfigure(1, weight=1)
        
        ctk.CTkLabel(
            step2,
            text="2. Source Name & Location",
            font=ctk.CTkFont(size=16, weight="bold")
        ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 10))
        
        ctk.CTkLabel(step2, text="Name:").grid(row=1, column=0, sticky="w", pady=5)
        self.name_entry = ctk.CTkEntry(step2, width=300)
        self.name_entry.grid(row=1, column=1, sticky="w", padx=10)
        
        ctk.CTkLabel(step2, text="Location:").grid(row=2, column=0, sticky="w", pady=5)
        self.location_label = ctk.CTkLabel(step2, text=str(self.config.sources_folder), text_color="#888888")
        self.location_label.grid(row=2, column=1, sticky="w", padx=10)
        
        ctk.CTkButton(step2, text="Choose Subfolder", width=130, command=self._choose_location).grid(row=2, column=2)
        
        # Step 3: Page range (for PDFs)
        self.step3 = ctk.CTkFrame(self)
        self.step3.grid(row=2, column=0, sticky="ew", padx=20, pady=10)
        self.step3.grid_columnconfigure(1, weight=1)
        
        ctk.CTkLabel(
            self.step3,
            text="3. Page Range",
            font=ctk.CTkFont(size=16, weight="bold")
        ).grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 10))
        
        ctk.CTkLabel(self.step3, text="From page:").grid(row=1, column=0, sticky="w", pady=5)
        self.start_page_entry = ctk.CTkEntry(self.step3, width=80)
        self.start_page_entry.grid(row=1, column=1, sticky="w", padx=10)
        self.start_page_entry.insert(0, "1")
        
        ctk.CTkLabel(self.step3, text="To page:").grid(row=1, column=2, sticky="w", padx=(20, 0))
        self.end_page_entry = ctk.CTkEntry(self.step3, width=80, placeholder_text="last")
        self.end_page_entry.grid(row=1, column=3, sticky="w", padx=10)
        
        # Buttons
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.grid(row=3, column=0, sticky="ew", padx=20, pady=20)
        
        ctk.CTkButton(btn_frame, text="Cancel", width=100, fg_color="transparent", border_width=1, command=self.destroy).pack(side="left")
        
        self.import_btn = ctk.CTkButton(
            btn_frame,
            text="Import Source",
            width=140,
            fg_color="#28a745",
            hover_color="#218838",
            state="disabled",
            command=self._import_source
        )
        self.import_btn.pack(side="right")
        
        self.source_file: Optional[Path] = None
        self.source_type: str = "pdf"
        self.dest_folder: Path = self.config.sources_folder
    
    def _select_pdf(self):
        filepath = filedialog.askopenfilename(
            title="Select PDF",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if filepath:
            self.source_file = Path(filepath)
            self.source_type = "pdf"
            self.source_path_label.configure(text=str(self.source_file))
            self.name_entry.delete(0, "end")
            self.name_entry.insert(0, self.source_file.stem)
            self.import_btn.configure(state="normal")
            
            # Get page count
            doc = fitz.open(filepath)
            self.end_page_entry.configure(placeholder_text=str(len(doc)))
            doc.close()
    
    def _select_png_folder(self):
        folder = filedialog.askdirectory(title="Select PNG Folder")
        if folder:
            self.source_file = Path(folder)
            self.source_type = "png_folder"
            self.source_path_label.configure(text=str(self.source_file))
            self.name_entry.delete(0, "end")
            self.name_entry.insert(0, self.source_file.name)
            self.import_btn.configure(state="normal")
            self.step3.grid_forget()  # Hide page range for PNG folders
    
    def _choose_location(self):
        folder = filedialog.askdirectory(
            title="Choose Location in Sources Folder",
            initialdir=str(self.config.sources_folder)
        )
        if folder:
            folder_path = Path(folder)
            # Ensure it's within sources folder
            try:
                folder_path.relative_to(self.config.sources_folder)
                self.dest_folder = folder_path
                self.location_label.configure(text=str(folder_path))
            except ValueError:
                messagebox.showerror("Error", "Please choose a location within your sources folder.")
    
    def _import_source(self):
        if not self.source_file:
            return
        
        name = self.name_entry.get().strip()
        if not name:
            messagebox.showerror("Error", "Please enter a source name.")
            return
        
        # Create source folder
        source_folder = self.dest_folder / name
        if source_folder.exists():
            messagebox.showerror("Error", f"A source named '{name}' already exists in this location.")
            return
        
        try:
            source_folder.mkdir(parents=True)
            
            if self.source_type == "pdf":
                # Copy PDF
                dest_pdf = source_folder / self.source_file.name
                shutil.copy2(self.source_file, dest_pdf)
                
                # Get page range
                start = int(self.start_page_entry.get() or "1")
                end_text = self.end_page_entry.get().strip()
                end = int(end_text) if end_text else None
                
                # Create metadata
                meta = {
                    "name": name,
                    "type": "pdf",
                    "original_file": str(self.source_file),
                    "page_range": {"start": start, "end": end},
                    "default_crop": {"left": 0, "right": 0, "top": 0, "bottom": 0},
                    "page_tags": {},
                    "page_crops": {},
                    "created": datetime.now().isoformat(),
                }
            else:
                # Copy PNG folder contents
                for png in self.source_file.glob("*.png"):
                    shutil.copy2(png, source_folder / png.name)
                for png in self.source_file.glob("*.PNG"):
                    shutil.copy2(png, source_folder / png.name)
                
                meta = {
                    "name": name,
                    "type": "png_folder",
                    "original_file": str(self.source_file),
                    "page_range": {"start": 1, "end": None},
                    "default_crop": {"left": 0, "right": 0, "top": 0, "bottom": 0},
                    "page_tags": {},
                    "page_crops": {},
                    "created": datetime.now().isoformat(),
                }
            
            save_json(source_folder / "source_meta.json", meta)
            
            self.callback()
            self.destroy()
            
            # Open the source for editing
            source = Source(source_folder)
            self.app.show_source_editor(source)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to import source: {e}")
            if source_folder.exists():
                shutil.rmtree(source_folder)


class SourceEditor(ctk.CTkFrame):
    """Edit a source - set crops, tag pages, export."""
    
    def __init__(self, parent, app, source: Source):
        super().__init__(parent)
        self.app = app
        self.source = source
        
        app.config.add_recent_source(source.path)
        
        self.current_page = source.meta.get("last_page", 1)
        self.page_images: list[Image.Image] = []
        self.total_pages = 0
        self.display_scale = source.meta.get("zoom", 0.5)
        
        self.margin_left = tk.DoubleVar(value=source.get_default_crop()["left"])
        self.margin_right = tk.DoubleVar(value=source.get_default_crop()["right"])
        self.margin_top = tk.DoubleVar(value=source.get_default_crop()["top"])
        self.margin_bottom = tk.DoubleVar(value=source.get_default_crop()["bottom"])
        
        self.show_crop_lines = tk.BooleanVar(value=True)
        self.per_page_mode = tk.BooleanVar(value=False)
        
        self.showing_average = False
        self.average_image: Optional[Image.Image] = None
        self.average_image_original: Optional[Image.Image] = None
        self.avg_contrast = tk.DoubleVar(value=1.5)
        self.avg_brightness = tk.DoubleVar(value=1.0)
        self.avg_sharpness = tk.DoubleVar(value=1.0)
        
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        
        self._setup_sidebar()
        self._setup_viewer()
        self._load_pages()
        self._load_page_margins()
        self._update_display()
    
    def _setup_sidebar(self):
        sidebar = ctk.CTkFrame(self, width=300)
        sidebar.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        sidebar.grid_propagate(False)
        
        # Header
        header = ctk.CTkFrame(sidebar, fg_color="transparent")
        header.pack(fill="x", padx=15, pady=15)
        
        ctk.CTkButton(
            header,
            text="< Back",
            width=70,
            command=self.app.show_source_browser
        ).pack(side="left")
        
        ctk.CTkLabel(
            header,
            text=self.source.name,
            font=ctk.CTkFont(size=16, weight="bold"),
            wraplength=180
        ).pack(side="left", padx=10)
        
        # Crop Margins
        ctk.CTkLabel(
            sidebar,
            text="Crop Margins",
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(pady=(10, 5), padx=15, anchor="w")
        
        self._create_margin_slider(sidebar, "Left", self.margin_left)
        self._create_margin_slider(sidebar, "Right", self.margin_right)
        self._create_margin_slider(sidebar, "Top", self.margin_top)
        self._create_margin_slider(sidebar, "Bottom", self.margin_bottom)
        
        ctk.CTkCheckBox(
            sidebar,
            text="Show crop lines",
            variable=self.show_crop_lines,
            command=self._update_display
        ).pack(pady=10, padx=15, anchor="w")
        
        ctk.CTkButton(
            sidebar,
            text="Save as Default Crop",
            command=self._save_default_crop
        ).pack(fill="x", padx=15, pady=5)
        
        # Separator
        ctk.CTkFrame(sidebar, height=2).pack(fill="x", padx=15, pady=15)
        
        # Actions
        ctk.CTkLabel(
            sidebar,
            text="Actions",
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(pady=(5, 8), padx=15, anchor="w")
        
        ctk.CTkButton(
            sidebar,
            text="Copy Current Page",
            command=self._copy_page
        ).pack(fill="x", padx=15, pady=2)
        
        ctk.CTkButton(
            sidebar,
            text="Auto-Detect Margins",
            command=self._auto_detect_margins
        ).pack(fill="x", padx=15, pady=2)
        
        ctk.CTkButton(
            sidebar,
            text="Export to PNG Folder",
            fg_color="#28a745",
            hover_color="#218838",
            command=self._export_source
        ).pack(fill="x", padx=15, pady=2)
        
        # Separator
        ctk.CTkFrame(sidebar, height=2).pack(fill="x", padx=15, pady=15)
        
        # Tags section
        ctk.CTkLabel(
            sidebar,
            text="Page Tags",
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(pady=(5, 5), padx=15, anchor="w")
        
        self.tags_label = ctk.CTkLabel(
            sidebar,
            text="No tags",
            text_color="#888888",
            wraplength=260
        )
        self.tags_label.pack(padx=15, anchor="w")
        
        tags_btn_frame = ctk.CTkFrame(sidebar, fg_color="transparent")
        tags_btn_frame.pack(fill="x", padx=15, pady=5)
        
        ctk.CTkButton(
            tags_btn_frame,
            text="Add Tag",
            width=80,
            command=self._add_tag
        ).pack(side="left", padx=(0, 5))
        
        ctk.CTkButton(
            tags_btn_frame,
            text="Remove Tag",
            width=90,
            command=self._remove_tag
        ).pack(side="left", padx=(0, 5))
        
        ctk.CTkButton(
            tags_btn_frame,
            text="Bulk Tag",
            width=80,
            command=self._bulk_tag
        ).pack(side="left")
        
        ctk.CTkButton(
            sidebar,
            text="AI Auto-Tag",
            fg_color="#6f42c1",
            hover_color="#5a32a3",
            command=self._ai_auto_tag
        ).pack(fill="x", padx=15, pady=(8, 0))
        
        self.quick_tags_frame = ctk.CTkFrame(sidebar, fg_color="transparent")
        self.quick_tags_frame.pack(fill="x", padx=15, pady=(5, 0))
        
        self.all_tags_label = ctk.CTkLabel(
            sidebar,
            text="",
            text_color="#666666",
            wraplength=260
        )
        self.all_tags_label.pack(padx=15, pady=(5, 0), anchor="w")
        
        # Status
        self.status_label = ctk.CTkLabel(sidebar, text="", wraplength=260)
        self.status_label.pack(padx=15, pady=10, anchor="w")
    
    def _create_margin_slider(self, parent, name: str, variable: tk.DoubleVar):
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.pack(fill="x", padx=15, pady=3)
        
        ctk.CTkLabel(frame, text=f"{name}:", width=50, anchor="w").pack(side="left")
        
        value_label = ctk.CTkLabel(frame, text=str(int(variable.get())), width=40)
        value_label.pack(side="right")
        
        slider = ctk.CTkSlider(
            frame,
            from_=0,
            to=500,
            variable=variable,
            command=lambda v, vl=value_label: self._on_margin_change(v, vl)
        )
        slider.pack(side="left", expand=True, fill="x", padx=5)
    
    def _on_margin_change(self, value: float, value_label: ctk.CTkLabel):
        value_label.configure(text=str(int(value)))
        self._update_display()
        
        if self.per_page_mode.get():
            self._auto_save_page_crop()
        else:
            self._auto_save_default_crop()
    
    def _auto_save_default_crop(self):
        crop = {
            "left": self.margin_left.get(),
            "right": self.margin_right.get(),
            "top": self.margin_top.get(),
            "bottom": self.margin_bottom.get()
        }
        self.source.set_default_crop(crop)
    
    def _setup_viewer(self):
        self.viewer = ctk.CTkFrame(self)
        self.viewer.grid(row=0, column=1, sticky="nsew", padx=(0, 10), pady=10)
        self.viewer.grid_columnconfigure(0, weight=1)
        self.viewer.grid_rowconfigure(2, weight=1)
        
        self.per_page_bar = ctk.CTkFrame(self.viewer)
        self.per_page_bar.grid(row=0, column=0, sticky="ew", padx=5, pady=5)
        
        self.per_page_toggle = ctk.CTkSwitch(
            self.per_page_bar,
            text="Per-Page Edit Mode",
            variable=self.per_page_mode,
            command=self._on_per_page_toggle
        )
        self.per_page_toggle.pack(side="left", padx=10)
        
        self.per_page_status = ctk.CTkLabel(
            self.per_page_bar,
            text="Global margins apply to all pages",
            text_color="#888888"
        )
        self.per_page_status.pack(side="left", padx=10)
        
        self.save_page_btn = ctk.CTkButton(
            self.per_page_bar,
            text="Save Page Override",
            width=140,
            command=self._save_page_crop,
            state="disabled"
        )
        self.save_page_btn.pack(side="right", padx=5)
        
        self.clear_page_btn = ctk.CTkButton(
            self.per_page_bar,
            text="Clear Override",
            width=120,
            fg_color="#dc3545",
            hover_color="#c82333",
            command=self._clear_page_crop,
            state="disabled"
        )
        self.clear_page_btn.pack(side="right", padx=5)
        
        self.avg_controls_frame = ctk.CTkFrame(self.viewer)
        
        ctk.CTkLabel(self.avg_controls_frame, text="Contrast:", width=70).pack(side="left", padx=(10, 5))
        self.contrast_slider = ctk.CTkSlider(
            self.avg_controls_frame, from_=0.5, to=3.0, width=100,
            variable=self.avg_contrast, command=lambda v: self._update_average_enhancements()
        )
        self.contrast_slider.pack(side="left", padx=5)
        
        ctk.CTkLabel(self.avg_controls_frame, text="Brightness:", width=80).pack(side="left", padx=(15, 5))
        self.brightness_slider = ctk.CTkSlider(
            self.avg_controls_frame, from_=0.5, to=2.0, width=100,
            variable=self.avg_brightness, command=lambda v: self._update_average_enhancements()
        )
        self.brightness_slider.pack(side="left", padx=5)
        
        ctk.CTkLabel(self.avg_controls_frame, text="Sharpness:", width=80).pack(side="left", padx=(15, 5))
        self.sharpness_slider = ctk.CTkSlider(
            self.avg_controls_frame, from_=0.0, to=3.0, width=100,
            variable=self.avg_sharpness, command=lambda v: self._update_average_enhancements()
        )
        self.sharpness_slider.pack(side="left", padx=5)
        
        ctk.CTkButton(
            self.avg_controls_frame, text="Reset Enhance", width=100,
            command=self._reset_average_enhancements
        ).pack(side="left", padx=15)
        
        self.exit_average_btn = ctk.CTkButton(
            self.avg_controls_frame,
            text="Exit Average View",
            width=140,
            fg_color="#ffc107",
            hover_color="#e0a800",
            text_color="#000000",
            command=self._exit_average_mode
        )
        self.exit_average_btn.pack(side="right", padx=10)
        
        nav = ctk.CTkFrame(self.viewer, fg_color="transparent")
        nav.grid(row=1, column=0, sticky="ew", padx=10, pady=10)
        
        self.prev_btn = ctk.CTkButton(nav, text="< Prev", width=80, command=self._prev_page)
        self.prev_btn.pack(side="left")
        
        self.page_label = ctk.CTkLabel(nav, text="Page 0 / 0")
        self.page_label.pack(side="left", expand=True)
        
        self.next_btn = ctk.CTkButton(nav, text="Next >", width=80, command=self._next_page)
        self.next_btn.pack(side="right")
        
        canvas_frame = ctk.CTkFrame(self.viewer)
        canvas_frame.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 10))
        canvas_frame.grid_columnconfigure(0, weight=1)
        canvas_frame.grid_rowconfigure(0, weight=1)
        
        self.canvas = tk.Canvas(canvas_frame, bg="#2b2b2b", highlightthickness=0)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        
        scrollbar = ctk.CTkScrollbar(canvas_frame, command=self.canvas.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.canvas.configure(yscrollcommand=scrollbar.set)
        
        zoom_frame = ctk.CTkFrame(self.viewer, fg_color="transparent")
        zoom_frame.grid(row=3, column=0, sticky="ew", padx=10, pady=(0, 10))
        
        ctk.CTkLabel(zoom_frame, text="Zoom:").pack(side="left", padx=(0, 10))
        
        self.zoom_slider = ctk.CTkSlider(
            zoom_frame,
            from_=0.25,
            to=2.0,
            command=self._on_zoom
        )
        self.zoom_slider.set(self.display_scale)
        self.zoom_slider.pack(side="left", expand=True, fill="x", padx=10)
        
        self.zoom_label = ctk.CTkLabel(zoom_frame, text=f"{int(self.display_scale * 100)}%", width=50)
        self.zoom_label.pack(side="right")
        
        self.canvas.bind("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind("<Button-4>", self._on_mousewheel)
        self.canvas.bind("<Button-5>", self._on_mousewheel)
        
        self._bind_keyboard_shortcuts()
    
    def _load_pages(self):
        self.page_images = []
        
        if self.source.source_type == "pdf":
            pdf_path = self.source.pdf_path
            if not pdf_path or not pdf_path.exists():
                return
            
            doc = fitz.open(str(pdf_path))
            zoom = 150 / 72
            mat = fitz.Matrix(zoom, zoom)
            
            for page_num in range(len(doc)):
                page = doc[page_num]
                pix = page.get_pixmap(matrix=mat)
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                self.page_images.append(img)
            
            doc.close()
        else:
            # PNG folder
            pngs = sorted(
                list(self.source.path.glob("*.png")) + list(self.source.path.glob("*.PNG")),
                key=lambda p: p.name.lower()
            )
            for png in pngs:
                if png.name != "source_meta.json":
                    img = Image.open(png)
                    self.page_images.append(img.copy())
                    img.close()
        
        self.total_pages = len(self.page_images)
        self.current_page = 1
    
    def _load_page_margins(self):
        if self.per_page_mode.get() and self.source.has_page_crop_override(self.current_page):
            crop = self.source.get_page_crop(self.current_page)
        else:
            crop = self.source.get_default_crop()
        
        self.margin_left.set(crop["left"])
        self.margin_right.set(crop["right"])
        self.margin_top.set(crop["top"])
        self.margin_bottom.set(crop["bottom"])
        
        if self.per_page_mode.get():
            self._update_per_page_indicator()
        
        self._update_tags_display()
        self._update_all_tags_display()
    
    def _update_tags_display(self):
        tags = self.source.get_page_tags(self.current_page)
        if tags:
            self.tags_label.configure(text=", ".join(tags))
        else:
            self.tags_label.configure(text="No tags")
    
    def _update_all_tags_display(self):
        all_tags = self.source.get_all_tags()
        if all_tags:
            self.all_tags_label.configure(text=f"All tags: {', '.join(sorted(all_tags))}")
        else:
            self.all_tags_label.configure(text="")
        
        self._update_quick_tags()
    
    def _update_quick_tags(self):
        for widget in self.quick_tags_frame.winfo_children():
            widget.destroy()
        
        all_tags = sorted(self.source.get_all_tags())
        if not all_tags:
            return
        
        current_tags = set(self.source.get_page_tags(self.current_page))
        
        row_frame = None
        for i, tag in enumerate(all_tags[:12]):
            if i % 4 == 0:
                row_frame = ctk.CTkFrame(self.quick_tags_frame, fg_color="transparent")
                row_frame.pack(fill="x", pady=1)
            
            is_active = tag in current_tags
            btn = ctk.CTkButton(
                row_frame,
                text=tag,
                width=60,
                height=24,
                font=ctk.CTkFont(size=11),
                fg_color="#28a745" if is_active else "transparent",
                border_width=1,
                command=lambda t=tag: self._toggle_quick_tag(t)
            )
            btn.pack(side="left", padx=2, pady=1)
    
    def _toggle_quick_tag(self, tag: str):
        current_tags = self.source.get_page_tags(self.current_page)
        if tag in current_tags:
            self.source.remove_page_tag(self.current_page, tag)
            self.status_label.configure(text=f"Removed '{tag}'")
        else:
            self.source.add_page_tag(self.current_page, tag)
            self.status_label.configure(text=f"Added '{tag}'")
        self._update_tags_display()
        self._update_quick_tags()
    
    def _update_display(self):
        if self.showing_average and self.average_image:
            img = self.average_image.copy()
        elif not self.page_images or self.current_page < 1 or self.current_page > self.total_pages:
            return
        else:
            img = self.page_images[self.current_page - 1].copy()
        
        if self.show_crop_lines.get():
            draw = ImageDraw.Draw(img)
            w, h = img.size
            
            left = int(self.margin_left.get())
            right = w - int(self.margin_right.get())
            top = int(self.margin_top.get())
            bottom = h - int(self.margin_bottom.get())
            
            line_color = "#FF0000"
            line_width = 2
            
            draw.line([(left, 0), (left, h)], fill=line_color, width=line_width)
            draw.line([(right, 0), (right, h)], fill=line_color, width=line_width)
            draw.line([(0, top), (w, top)], fill=line_color, width=line_width)
            draw.line([(0, bottom), (w, bottom)], fill=line_color, width=line_width)
            
            # Dim outside areas
            overlay = Image.new("RGBA", img.size, (0, 0, 0, 0))
            overlay_draw = ImageDraw.Draw(overlay)
            dim_color = (0, 0, 0, 100)
            
            if left > 0:
                overlay_draw.rectangle([(0, 0), (left, h)], fill=dim_color)
            if right < w:
                overlay_draw.rectangle([(right, 0), (w, h)], fill=dim_color)
            if top > 0:
                overlay_draw.rectangle([(left, 0), (right, top)], fill=dim_color)
            if bottom < h:
                overlay_draw.rectangle([(left, bottom), (right, h)], fill=dim_color)
            
            img = img.convert("RGBA")
            img = Image.alpha_composite(img, overlay)
            img = img.convert("RGB")
        
        # Apply zoom
        if self.display_scale != 1.0:
            new_size = (int(img.width * self.display_scale), int(img.height * self.display_scale))
            img = img.resize(new_size, Image.Resampling.LANCZOS)
        
        self.tk_image = ImageTk.PhotoImage(img)
        
        self.canvas.delete("all")
        
        canvas_width = self.canvas.winfo_width()
        canvas_height = self.canvas.winfo_height()
        
        x_offset = max(0, (canvas_width - img.width) // 2)
        y_offset = max(0, (canvas_height - img.height) // 2)
        
        self.canvas.create_image(x_offset, y_offset, anchor="nw", image=self.tk_image)
        self.canvas.configure(scrollregion=(0, 0, max(canvas_width, img.width), max(canvas_height, img.height)))
        
        if self.showing_average:
            start, end = self.source.get_page_range()
            page_count = min(end, self.total_pages) - start + 1
            self.page_label.configure(text=f"AVERAGE VIEW ({page_count} pages)")
        else:
            start, end = self.source.get_page_range()
            override_indicator = " [CUSTOM CROP]" if self.source.has_page_crop_override(self.current_page) else ""
            self.page_label.configure(text=f"Page {self.current_page} / {self.total_pages} (range: {start}-{end}){override_indicator}")
    
    def _prev_page(self):
        if self.current_page > 1:
            self.current_page -= 1
            self._save_position()
            self._load_page_margins()
            self._update_display()
    
    def _next_page(self):
        if self.current_page < self.total_pages:
            self.current_page += 1
            self._save_position()
            self._load_page_margins()
            self._update_display()
    
    def _save_position(self):
        self.source.meta["last_page"] = self.current_page
        self.source.meta["zoom"] = self.display_scale
        self.source.save_meta()
    
    def _on_zoom(self, value: float):
        self.display_scale = value
        self.zoom_label.configure(text=f"{int(value * 100)}%")
        self._save_position()
        self._update_display()
    
    def _on_mousewheel(self, event):
        if event.num == 4 or event.delta > 0:
            self.canvas.yview_scroll(-1, "units")
        elif event.num == 5 or event.delta < 0:
            self.canvas.yview_scroll(1, "units")
    
    def _bind_keyboard_shortcuts(self):
        self.command_buffer = ""
        
        self.bind("<Left>", lambda e: self._prev_page())
        self.bind("<Right>", lambda e: self._next_page())
        self.bind("<j>", lambda e: self._next_page())
        self.bind("<k>", lambda e: self._prev_page())
        self.bind("<g>", lambda e: self._go_to_page(1))
        self.bind("<G>", lambda e: self._go_to_page(self.total_pages))
        self.bind("<colon>", lambda e: self._open_command_input())
        self.bind("<question>", lambda e: self._show_keybind_help())
        self.bind("<c>", lambda e: self._copy_page())
        self.bind("<Key>", self._handle_vim_number)
        
        self.focus_set()
    
    def _handle_vim_number(self, event):
        if event.char.isdigit():
            self.command_buffer += event.char
            self.after(1500, self._clear_command_buffer)
            return
        
        if event.keysym == "Return" and self.command_buffer:
            try:
                num = int(self.command_buffer)
                self._go_to_page(num)
            except ValueError:
                pass
            self.command_buffer = ""
    
    def _clear_command_buffer(self):
        self.command_buffer = ""
    
    def _go_to_page(self, page_num: int):
        if not self.page_images:
            return
        page_num = max(1, min(page_num, self.total_pages))
        self.current_page = page_num
        self._load_page_margins()
        self._update_display()
    
    def _open_command_input(self):
        if not self.page_images:
            return
        CommandDialog(self, self._execute_command)
    
    def _execute_command(self, cmd: str):
        cmd = cmd.strip()
        if not cmd:
            return
        
        try:
            page_num = int(cmd)
            if 1 <= page_num <= self.total_pages:
                self._go_to_page(page_num)
            else:
                self.status_label.configure(text=f"Page {page_num} out of range (1-{self.total_pages})")
        except ValueError:
            self.status_label.configure(text=f"Invalid command: {cmd}")
    
    def _show_keybind_help(self):
        KeybindHelpDialog(self)
    
    def _on_per_page_toggle(self):
        if self.per_page_mode.get():
            self.save_page_btn.configure(state="normal")
            self.clear_page_btn.configure(state="normal")
            self._load_page_margins()
            self._update_per_page_indicator()
        else:
            self.save_page_btn.configure(state="disabled")
            self.clear_page_btn.configure(state="disabled")
            self.per_page_status.configure(text="Global margins apply to all pages", text_color="#888888")
    
    def _save_default_crop(self):
        crop = {
            "left": self.margin_left.get(),
            "right": self.margin_right.get(),
            "top": self.margin_top.get(),
            "bottom": self.margin_bottom.get()
        }
        self.source.set_default_crop(crop)
        self.status_label.configure(text="Default crop saved")
    
    def _save_page_crop(self):
        crop = {
            "left": self.margin_left.get(),
            "right": self.margin_right.get(),
            "top": self.margin_top.get(),
            "bottom": self.margin_bottom.get()
        }
        self.source.set_page_crop(self.current_page, crop)
        self.per_page_status.configure(text=f"Page {self.current_page} has custom crop")
        self.status_label.configure(text=f"Saved crop for page {self.current_page}")
        self._update_display()
    
    def _auto_save_page_crop(self):
        crop = {
            "left": self.margin_left.get(),
            "right": self.margin_right.get(),
            "top": self.margin_top.get(),
            "bottom": self.margin_bottom.get()
        }
        self.source.set_page_crop(self.current_page, crop)
        self._update_per_page_indicator()
    
    def _update_per_page_indicator(self):
        if self.source.has_page_crop_override(self.current_page):
            self.per_page_status.configure(
                text=f"CUSTOM CROP - Page {self.current_page}",
                text_color="#28a745"
            )
        else:
            self.per_page_status.configure(
                text=f"Page {self.current_page} uses default crop",
                text_color="#888888"
            )
    
    def _clear_page_crop(self):
        if self.source.has_page_crop_override(self.current_page):
            self.source.clear_page_crop(self.current_page)
            self.status_label.configure(text=f"Cleared override for page {self.current_page}")
            self._load_page_margins()
            self._update_per_page_indicator()
            self._update_display()
        else:
            self.status_label.configure(text=f"Page {self.current_page} has no override")
    
    def _add_tag(self):
        tag = simpledialog.askstring("Add Tag", f"Enter tag for page {self.current_page}:", parent=self)
        if tag:
            tag = tag.strip()
            if tag:
                self.source.add_page_tag(self.current_page, tag)
                self._update_tags_display()
                self._update_all_tags_display()
                self.status_label.configure(text=f"Added tag '{tag}' to page {self.current_page}")
    
    def _remove_tag(self):
        tags = self.source.get_page_tags(self.current_page)
        if not tags:
            messagebox.showinfo("No Tags", "This page has no tags to remove.")
            return
        
        tag = simpledialog.askstring(
            "Remove Tag",
            f"Current tags: {', '.join(tags)}\nEnter tag to remove:",
            parent=self
        )
        if tag:
            tag = tag.strip()
            if tag in tags:
                self.source.remove_page_tag(self.current_page, tag)
                self._update_tags_display()
                self._update_all_tags_display()
                self.status_label.configure(text=f"Removed tag '{tag}' from page {self.current_page}")
            else:
                messagebox.showinfo("Not Found", f"Tag '{tag}' not found on this page.")
    
    def _bulk_tag(self):
        BulkTagDialog(self, self.source, self._update_tags_display, self._update_all_tags_display)
    
    def _ai_auto_tag(self):
        if not self.page_images:
            messagebox.showinfo("No Pages", "No pages loaded.")
            return
        AIAutoTagDialog(self, self.source, self.page_images, self._update_tags_display, self._update_all_tags_display, self.app.config)
    
    def _copy_page(self):
        if not self.page_images:
            return
        
        try:
            img = self._get_cropped_image(self.current_page)
            copy_image_to_clipboard(img)
            
            self.app.config.add_history({
                "action": "copy",
                "source": self.source.name,
                "source_path": str(self.source.path),
                "page": self.current_page
            })
            
            self.status_label.configure(text=f"Copied page {self.current_page} to clipboard")
            
        except Exception as e:
            messagebox.showerror("Error", f"Copy failed: {e}")
    
    def _get_cropped_image(self, page_num: int) -> Image.Image:
        img = self.page_images[page_num - 1]
        w, h = img.size
        
        crop = self.source.get_page_crop(page_num)
        left = int(crop["left"])
        right = int(crop["right"])
        top = int(crop["top"])
        bottom = int(crop["bottom"])
        
        return img.crop((left, top, w - right, h - bottom))
    
    def _auto_detect_margins(self):
        if not self.page_images:
            return
        
        self.average_image_original = self._compute_average_image()
        
        if not self.average_image_original:
            self.status_label.configure(text="Could not compute average image")
            return
        
        self.showing_average = True
        self.avg_controls_frame.grid(row=0, column=0, sticky="ew", padx=5, pady=5)
        self.per_page_bar.grid_forget()
        self._update_average_enhancements()
        
        start, end = self.source.get_page_range()
        page_count = min(end, self.total_pages) - start + 1
        self.status_label.configure(text=f"Showing average of {page_count} pages - adjust margins, then Exit")
    
    def _compute_average_image(self) -> Optional[Image.Image]:
        if not self.page_images:
            return None
        
        import numpy as np
        
        start, end = self.source.get_page_range()
        
        arrays = []
        for i in range(start - 1, min(end, self.total_pages)):
            img = self.page_images[i]
            arrays.append(np.array(img, dtype=np.float32))
        
        if not arrays:
            return None
        
        avg_array = np.mean(arrays, axis=0).astype(np.uint8)
        return Image.fromarray(avg_array)
    
    def _update_average_enhancements(self):
        if not self.average_image_original:
            return
        
        img = self.average_image_original.copy()
        
        contrast = self.avg_contrast.get()
        if contrast != 1.0:
            img = ImageEnhance.Contrast(img).enhance(contrast)
        
        brightness = self.avg_brightness.get()
        if brightness != 1.0:
            img = ImageEnhance.Brightness(img).enhance(brightness)
        
        sharpness = self.avg_sharpness.get()
        if sharpness != 1.0:
            img = ImageEnhance.Sharpness(img).enhance(sharpness)
        
        self.average_image = img
        self._update_display()
    
    def _reset_average_enhancements(self):
        self.avg_contrast.set(1.5)
        self.avg_brightness.set(1.0)
        self.avg_sharpness.set(1.0)
        self._update_average_enhancements()
    
    def _exit_average_mode(self):
        self.showing_average = False
        self.average_image = None
        self.average_image_original = None
        self.avg_controls_frame.grid_forget()
        self.per_page_bar.grid(row=0, column=0, sticky="ew", padx=5, pady=5)
        self._update_display()
        self.status_label.configure(text="")
    
    def _export_source(self):
        if not self.page_images:
            return
        
        initial_dir = self.app.config.last_export_folder
        output_dir = filedialog.askdirectory(
            title="Select Output Folder",
            initialdir=str(initial_dir) if initial_dir else None
        )
        if not output_dir:
            return
        
        output_path = Path(output_dir)
        self.app.config.last_export_folder = output_path
        
        start, end = self.source.get_page_range()
        end = min(end or self.total_pages, self.total_pages)
        total_to_export = end - start + 1
        
        progress = ProgressDialog(self, "Exporting", total_to_export)
        
        try:
            exported = 0
            for i in range(start - 1, end):
                page_num = i + 1
                cropped = self._get_cropped_image(page_num)
                output_file = output_path / f"p{page_num}.png"
                cropped.save(output_file, "PNG")
                exported += 1
                progress.update_progress(exported, f"Exporting page {page_num}...")
                self.update()
            
            progress.destroy()
            
            self.app.config.add_history({
                "action": "export",
                "source": self.source.name,
                "source_path": str(self.source.path),
                "pages": f"{start}-{end}",
                "output": str(output_path),
                "count": exported
            })
            
            self.status_label.configure(text=f"Exported {exported} pages to {output_dir}")
            messagebox.showinfo("Export Complete", f"Exported {exported} cropped pages to:\n{output_dir}")
            
        except Exception as e:
            progress.destroy()
            messagebox.showerror("Error", f"Export failed: {e}")


class BulkTagDialog(ctk.CTkToplevel):
    """Dialog to bulk tag pages."""
    
    def __init__(self, parent, source: Source, update_callback, update_all_callback):
        super().__init__(parent)
        self.source = source
        self.update_callback = update_callback
        self.update_all_callback = update_all_callback
        
        self.title("Bulk Tag Pages")
        self.geometry("400x300")
        self.transient(parent)
        self.grab_set()
        
        self.grid_columnconfigure(0, weight=1)
        
        ctk.CTkLabel(
            self,
            text="Bulk Tag Pages",
            font=ctk.CTkFont(size=18, weight="bold")
        ).grid(row=0, column=0, pady=20)
        
        # Page range
        range_frame = ctk.CTkFrame(self, fg_color="transparent")
        range_frame.grid(row=1, column=0, pady=10, padx=20, sticky="ew")
        
        ctk.CTkLabel(range_frame, text="Pages:").pack(side="left")
        
        self.start_entry = ctk.CTkEntry(range_frame, width=60)
        self.start_entry.pack(side="left", padx=5)
        self.start_entry.insert(0, "1")
        
        ctk.CTkLabel(range_frame, text="to").pack(side="left", padx=5)
        
        self.end_entry = ctk.CTkEntry(range_frame, width=60, placeholder_text="last")
        self.end_entry.pack(side="left", padx=5)
        
        # Tag input
        tag_frame = ctk.CTkFrame(self, fg_color="transparent")
        tag_frame.grid(row=2, column=0, pady=10, padx=20, sticky="ew")
        
        ctk.CTkLabel(tag_frame, text="Tag:").pack(side="left")
        
        self.tag_entry = ctk.CTkEntry(tag_frame, width=200)
        self.tag_entry.pack(side="left", padx=10)
        
        # Action
        action_frame = ctk.CTkFrame(self, fg_color="transparent")
        action_frame.grid(row=3, column=0, pady=10)
        
        self.action_var = tk.StringVar(value="add")
        
        ctk.CTkRadioButton(action_frame, text="Add tag", variable=self.action_var, value="add").pack(side="left", padx=10)
        ctk.CTkRadioButton(action_frame, text="Remove tag", variable=self.action_var, value="remove").pack(side="left", padx=10)
        
        # Buttons
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.grid(row=4, column=0, pady=20)
        
        ctk.CTkButton(btn_frame, text="Cancel", width=80, fg_color="transparent", border_width=1, command=self.destroy).pack(side="left", padx=10)
        ctk.CTkButton(btn_frame, text="Apply", width=80, command=self._apply).pack(side="left", padx=10)
    
    def _apply(self):
        tag = self.tag_entry.get().strip()
        if not tag:
            messagebox.showerror("Error", "Please enter a tag.")
            return
        
        try:
            start = int(self.start_entry.get() or "1")
            end_text = self.end_entry.get().strip()
            end = int(end_text) if end_text else self.source.get_page_count()
        except ValueError:
            messagebox.showerror("Error", "Invalid page range.")
            return
        
        action = self.action_var.get()
        count = 0
        
        for page in range(start, end + 1):
            if action == "add":
                self.source.add_page_tag(page, tag)
                count += 1
            else:
                if tag in self.source.get_page_tags(page):
                    self.source.remove_page_tag(page, tag)
                    count += 1
        
        self.update_callback()
        self.update_all_callback()
        
        action_word = "Added" if action == "add" else "Removed"
        messagebox.showinfo("Done", f"{action_word} tag '{tag}' on {count} pages.")
        self.destroy()


class AIAutoTagDialog(ctk.CTkToplevel):
    def __init__(self, parent, source: Source, page_images: list, update_callback, update_all_callback, config: AppConfig):
        super().__init__(parent)
        self.source = source
        self.page_images = page_images
        self.update_callback = update_callback
        self.update_all_callback = update_all_callback
        self.config = config
        
        self.title("AI Auto-Tag")
        self.geometry("1200x750")
        self.transient(parent)
        self.grab_set()
        
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=2)
        self.grid_columnconfigure(2, weight=1)
        self.grid_rowconfigure(1, weight=1)
        
        self.tag_definitions: list[dict] = []
        self.results: list[dict] = []
        self.result_frames: dict[int, ctk.CTkFrame] = {}
        self.is_running = False
        self.sort_mode = tk.StringVar(value="page")
        self.selected_result_page: Optional[int] = None
        
        self._setup_header()
        self._setup_left_panel()
        self._setup_middle_panel()
        self._setup_preview_panel()
    
    def _setup_header(self):
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.grid(row=0, column=0, columnspan=3, sticky="ew", padx=20, pady=15)
        
        ctk.CTkLabel(
            header,
            text="AI Auto-Tag",
            font=ctk.CTkFont(size=20, weight="bold")
        ).pack(side="left")
        
        ctk.CTkLabel(
            header,
            text=f"Source: {self.source.name} ({len(self.page_images)} pages)",
            text_color="#888888"
        ).pack(side="left", padx=20)
        
        settings_frame = ctk.CTkFrame(header, fg_color="transparent")
        settings_frame.pack(side="right")
        
        ctk.CTkLabel(settings_frame, text="Model:").pack(side="left", padx=(0, 5))
        self.model_var = tk.StringVar(value=VISION_MODELS[0][1])
        model_menu = ctk.CTkOptionMenu(
            settings_frame,
            values=[m[0] for m in VISION_MODELS],
            command=self._on_model_change,
            width=200
        )
        model_menu.pack(side="left", padx=(0, 15))
        
        ctk.CTkLabel(settings_frame, text="Workers:").pack(side="left", padx=(0, 5))
        self.workers_var = tk.IntVar(value=10)
        workers_slider = ctk.CTkSlider(settings_frame, from_=1, to=20, number_of_steps=19, variable=self.workers_var, width=100)
        workers_slider.pack(side="left", padx=(0, 5))
        self.workers_label = ctk.CTkLabel(settings_frame, text="10", width=30)
        self.workers_label.pack(side="left")
        workers_slider.configure(command=lambda v: self.workers_label.configure(text=str(int(v))))
        
        ctk.CTkButton(
            settings_frame, text="API Key", width=80,
            fg_color="#6c757d", hover_color="#5a6268",
            command=self._set_api_key
        ).pack(side="left", padx=(15, 0))
    
    def _on_model_change(self, display_name: str):
        for name, model_id in VISION_MODELS:
            if name == display_name:
                self.model_var.set(model_id)
                break
    
    def _set_api_key(self):
        current = self.config.api_key or ""
        masked = current[:10] + "..." if len(current) > 10 else current
        
        key = simpledialog.askstring(
            "OpenRouter API Key",
            f"Current: {masked if current else '(not set)'}\n\nEnter your OpenRouter API key:\n(Get one at https://openrouter.ai/keys)",
            parent=self
        )
        if key is not None:
            self.config.api_key = key.strip()
            messagebox.showinfo("Saved", "API key saved.", parent=self)
    
    def _setup_left_panel(self):
        left = ctk.CTkFrame(self)
        left.grid(row=1, column=0, sticky="nsew", padx=(20, 10), pady=(0, 20))
        left.grid_rowconfigure(2, weight=1)
        left.grid_columnconfigure(0, weight=1)
        
        saved_defs = self.source.get_tag_definitions()
        if saved_defs:
            saved_frame = ctk.CTkFrame(left)
            saved_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 5))
            
            ctk.CTkLabel(saved_frame, text="Saved Definitions:", font=ctk.CTkFont(size=12, weight="bold")).pack(anchor="w", padx=10, pady=(5, 5))
            
            btns_frame = ctk.CTkFrame(saved_frame, fg_color="transparent")
            btns_frame.pack(fill="x", padx=10, pady=(0, 10))
            
            for td in saved_defs:
                btn = ctk.CTkButton(
                    btns_frame,
                    text=td["name"],
                    width=len(td["name"]) * 10 + 20,
                    height=28,
                    fg_color="#6f42c1",
                    hover_color="#5a32a3",
                    command=lambda t=td: self._add_tag_definition(t["name"], t["description"])
                )
                btn.pack(side="left", padx=2, pady=2)
            
            ctk.CTkButton(
                btns_frame,
                text="Load All",
                width=70,
                height=28,
                fg_color="#17a2b8",
                hover_color="#138496",
                command=self._load_all_saved
            ).pack(side="right", padx=2)
        
        header_frame = ctk.CTkFrame(left, fg_color="transparent")
        header_frame.grid(row=1, column=0, sticky="ew", padx=15, pady=(10, 5))
        
        ctk.CTkLabel(header_frame, text="Tag Definitions", font=ctk.CTkFont(size=14, weight="bold")).pack(side="left")
        
        ctk.CTkButton(
            header_frame, text="Save", width=60, height=26,
            fg_color="#28a745", hover_color="#218838",
            command=self._save_definitions
        ).pack(side="right")
        
        self.tags_scroll = ctk.CTkScrollableFrame(left)
        self.tags_scroll.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 10))
        
        add_frame = ctk.CTkFrame(left, fg_color="transparent")
        add_frame.grid(row=3, column=0, sticky="ew", padx=10, pady=(0, 10))
        
        ctk.CTkButton(add_frame, text="+ Add Tag Definition", command=self._add_tag_definition).pack(fill="x")
    
    def _load_all_saved(self):
        for td in self.tag_definitions:
            if td["frame"].winfo_exists():
                td["frame"].destroy()
        self.tag_definitions = []
        
        for saved in self.source.get_tag_definitions():
            self._add_tag_definition(saved["name"], saved["description"])
    
    def _save_definitions(self):
        defs = self._get_tag_definitions()
        if not defs:
            messagebox.showwarning("No Definitions", "Add at least one tag definition to save.")
            return
        
        self.source.set_tag_definitions(defs)
        messagebox.showinfo("Saved", f"Saved {len(defs)} tag definitions to source.")
    
    def _add_tag_definition(self, name: str = "", description: str = ""):
        frame = ctk.CTkFrame(self.tags_scroll)
        frame.pack(fill="x", pady=5, padx=5)
        
        top_row = ctk.CTkFrame(frame, fg_color="transparent")
        top_row.pack(fill="x", padx=10, pady=(10, 5))
        
        ctk.CTkLabel(top_row, text="Tag name:", width=80).pack(side="left")
        name_entry = ctk.CTkEntry(top_row, width=150, placeholder_text="e.g., math")
        name_entry.pack(side="left", padx=5)
        if name:
            name_entry.insert(0, name)
        
        ctk.CTkButton(
            top_row, text="X", width=30, fg_color="#dc3545", hover_color="#c82333",
            command=lambda f=frame: self._remove_tag_definition(f)
        ).pack(side="right")
        
        ctk.CTkLabel(frame, text="Description (what makes a page match this tag):", anchor="w").pack(
            fill="x", padx=10, pady=(5, 2)
        )
        desc_entry = ctk.CTkTextbox(frame, height=60)
        desc_entry.pack(fill="x", padx=10, pady=(0, 10))
        if description:
            desc_entry.insert("0.0", description)
        else:
            desc_entry.insert("0.0", "Pages that contain...")
        
        self.tag_definitions.append({
            "frame": frame,
            "name_entry": name_entry,
            "desc_entry": desc_entry
        })
    
    def _remove_tag_definition(self, frame):
        frame.destroy()
        self.tag_definitions = [td for td in self.tag_definitions if td["frame"].winfo_exists()]
    
    def _setup_middle_panel(self):
        middle = ctk.CTkFrame(self)
        middle.grid(row=1, column=1, sticky="nsew", padx=10, pady=(0, 20))
        middle.grid_rowconfigure(2, weight=1)
        middle.grid_columnconfigure(0, weight=1)
        
        top_bar = ctk.CTkFrame(middle, fg_color="transparent")
        top_bar.grid(row=0, column=0, sticky="ew", padx=15, pady=(15, 5))
        
        ctk.CTkLabel(top_bar, text="Results", font=ctk.CTkFont(size=14, weight="bold")).pack(side="left")
        
        self.progress_label = ctk.CTkLabel(top_bar, text="", text_color="#888888")
        self.progress_label.pack(side="left", padx=20)
        
        self.progress_bar = ctk.CTkProgressBar(top_bar, width=150)
        self.progress_bar.pack(side="left", padx=10)
        self.progress_bar.set(0)
        
        sort_bar = ctk.CTkFrame(middle, fg_color="transparent")
        sort_bar.grid(row=1, column=0, sticky="ew", padx=15, pady=(5, 5))
        
        ctk.CTkLabel(sort_bar, text="Sort:", text_color="#888888").pack(side="left")
        ctk.CTkRadioButton(sort_bar, text="Page #", variable=self.sort_mode, value="page", command=self._resort_results).pack(side="left", padx=5)
        ctk.CTkRadioButton(sort_bar, text="Tagged First", variable=self.sort_mode, value="tagged", command=self._resort_results).pack(side="left", padx=5)
        ctk.CTkRadioButton(sort_bar, text="Untagged First", variable=self.sort_mode, value="untagged", command=self._resort_results).pack(side="left", padx=5)
        
        self.results_scroll = ctk.CTkScrollableFrame(middle)
        self.results_scroll.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 10))
        
        btn_frame = ctk.CTkFrame(middle, fg_color="transparent")
        btn_frame.grid(row=3, column=0, sticky="ew", padx=10, pady=(0, 15))
        
        self.trial_btn = ctk.CTkButton(
            btn_frame, text="Trial (20)", width=100,
            fg_color="#17a2b8", hover_color="#138496",
            command=self._run_trial
        )
        self.trial_btn.pack(side="left", padx=3)
        
        self.run_btn = ctk.CTkButton(
            btn_frame, text="Run All", width=80,
            fg_color="#28a745", hover_color="#218838",
            command=self._run_all
        )
        self.run_btn.pack(side="left", padx=3)
        
        self.apply_btn = ctk.CTkButton(
            btn_frame, text="Apply Tags", width=90,
            fg_color="#6f42c1", hover_color="#5a32a3",
            command=self._apply_tags,
            state="disabled"
        )
        self.apply_btn.pack(side="left", padx=3)
        
        ctk.CTkButton(
            btn_frame, text="Close", width=60,
            fg_color="transparent", border_width=1,
            command=self.destroy
        ).pack(side="right", padx=3)
    
    def _setup_preview_panel(self):
        preview = ctk.CTkFrame(self)
        preview.grid(row=1, column=2, sticky="nsew", padx=(10, 20), pady=(0, 20))
        preview.grid_rowconfigure(1, weight=1)
        preview.grid_columnconfigure(0, weight=1)
        
        ctk.CTkLabel(preview, text="Page Preview", font=ctk.CTkFont(size=14, weight="bold")).grid(
            row=0, column=0, sticky="w", padx=15, pady=(15, 10)
        )
        
        self.preview_canvas = tk.Canvas(preview, bg="#2b2b2b", highlightthickness=0)
        self.preview_canvas.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        
        self.preview_info = ctk.CTkLabel(preview, text="Click a result to preview", text_color="#888888")
        self.preview_info.grid(row=2, column=0, pady=(0, 10))
    
    def _resort_results(self):
        if not self.results:
            return
        
        for widget in self.results_scroll.winfo_children():
            widget.destroy()
        self.result_frames.clear()
        
        sorted_results = self.results.copy()
        mode = self.sort_mode.get()
        
        if mode == "page":
            sorted_results.sort(key=lambda r: r["page"])
        elif mode == "tagged":
            sorted_results.sort(key=lambda r: (0 if r.get("tags") else 1, r["page"]))
        elif mode == "untagged":
            sorted_results.sort(key=lambda r: (1 if r.get("tags") else 0, r["page"]))
        
        for result in sorted_results:
            self._add_result_row(result)
    
    def _get_tag_definitions(self) -> list[dict]:
        defs = []
        for td in self.tag_definitions:
            if not td["frame"].winfo_exists():
                continue
            name = td["name_entry"].get().strip()
            desc = td["desc_entry"].get("0.0", "end").strip()
            if name and desc:
                defs.append({"name": name, "description": desc})
        return defs
    
    def _image_to_base64(self, img: Image.Image, max_size: int = 1024) -> str:
        img_copy = img.copy()
        img_copy.thumbnail((max_size, max_size), Image.Resampling.LANCZOS)
        
        buffer = io.BytesIO()
        img_copy.save(buffer, format="JPEG", quality=85)
        return base64.b64encode(buffer.getvalue()).decode("utf-8")
    
    def _call_vision_api(self, img: Image.Image, tag_defs: list[dict], page_num: int) -> dict:
        tags_prompt = "\n".join([f"- **{td['name']}**: {td['description']}" for td in tag_defs])
        
        prompt = f"""Analyze this page image and determine which tags apply.

Tags to check:
{tags_prompt}

Respond with ONLY a JSON object in this exact format:
{{"tags": ["tag1", "tag2"], "reasoning": "brief explanation"}}

Only include tags that clearly match. If no tags match, return {{"tags": [], "reasoning": "explanation"}}"""

        img_base64 = self._image_to_base64(img)
        
        payload = {
            "model": self.model_var.get(),
            "messages": [
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": prompt},
                        {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{img_base64}"}}
                    ]
                }
            ],
            "max_tokens": 500,
            "temperature": 0.1
        }
        
        headers = {
            "Authorization": f"Bearer {self.config.api_key}",
            "Content-Type": "application/json",
            "HTTP-Referer": "https://pdf-crop-tool.local",
            "X-Title": "PDF Crop Tool"
        }
        
        try:
            req = urllib.request.Request(
                "https://openrouter.ai/api/v1/chat/completions",
                data=json.dumps(payload).encode("utf-8"),
                headers=headers,
                method="POST"
            )
            
            with urllib.request.urlopen(req, timeout=60) as response:
                result = json.loads(response.read().decode("utf-8"))
            
            content = result["choices"][0]["message"]["content"]
            
            json_start = content.find("{")
            json_end = content.rfind("}") + 1
            if json_start >= 0 and json_end > json_start:
                parsed = json.loads(content[json_start:json_end])
                return {
                    "page": page_num,
                    "tags": parsed.get("tags", []),
                    "reasoning": parsed.get("reasoning", ""),
                    "success": True
                }
            
            return {"page": page_num, "tags": [], "reasoning": "Failed to parse response", "success": False}
            
        except Exception as e:
            return {"page": page_num, "tags": [], "reasoning": str(e), "success": False}
    
    def _process_pages(self, page_indices: list[int], tag_defs: list[dict]):
        self.is_running = True
        self.results = []
        total = len(page_indices)
        completed = 0
        
        for widget in self.results_scroll.winfo_children():
            widget.destroy()
        self.result_frames.clear()
        
        self.progress_bar.set(0)
        self.progress_label.configure(text=f"0/{total}")
        self.trial_btn.configure(state="disabled")
        self.run_btn.configure(state="disabled")
        
        results_lock = threading.Lock()
        
        def process_page(page_idx: int) -> dict:
            img = self.page_images[page_idx]
            return self._call_vision_api(img, tag_defs, page_idx + 1)
        
        def update_ui(result: dict):
            nonlocal completed
            completed += 1
            
            with results_lock:
                self.results.append(result)
            
            self.after(0, lambda: self._add_result_row(result))
            self.after(0, lambda: self.progress_bar.set(completed / total))
            self.after(0, lambda: self.progress_label.configure(text=f"{completed}/{total}"))
        
        def run_workers():
            workers = self.workers_var.get()
            with ThreadPoolExecutor(max_workers=workers) as executor:
                futures = {executor.submit(process_page, idx): idx for idx in page_indices}
                for future in as_completed(futures):
                    if not self.is_running:
                        break
                    try:
                        result = future.result()
                        update_ui(result)
                    except Exception as e:
                        update_ui({"page": futures[future] + 1, "tags": [], "reasoning": str(e), "success": False})
            
            self.after(0, self._on_processing_complete)
        
        thread = threading.Thread(target=run_workers, daemon=True)
        thread.start()
    
    def _add_result_row(self, result: dict):
        page_num = result["page"]
        tags = result.get("tags", [])
        reasoning = result.get("reasoning", "")
        success = result.get("success", False)
        
        frame = ctk.CTkFrame(self.results_scroll, cursor="hand2")
        frame.pack(fill="x", pady=2, padx=5)
        frame.page_num = page_num
        self.result_frames[page_num] = frame
        
        color = "#28a745" if success and tags else "#888888" if success else "#dc3545"
        
        page_label = ctk.CTkLabel(frame, text=f"Page {page_num}:", width=60, anchor="w")
        page_label.pack(side="left", padx=(10, 5), pady=5)
        
        tags_text = ", ".join(tags) if tags else "(no tags)"
        tags_label = ctk.CTkLabel(frame, text=tags_text, text_color=color, anchor="w", width=120)
        tags_label.pack(side="left", padx=5, pady=5)
        
        reason_label = ctk.CTkLabel(
            frame, text=reasoning[:60] + "..." if len(reasoning) > 60 else reasoning,
            text_color="#888888", anchor="w", font=ctk.CTkFont(size=11)
        )
        reason_label.pack(side="left", padx=5, pady=5, fill="x", expand=True)
        
        for widget in [frame, page_label, tags_label, reason_label]:
            widget.bind("<Button-1>", lambda e, pn=page_num: self._select_result(pn))
    
    def _select_result(self, page_num: int):
        if self.selected_result_page and self.selected_result_page in self.result_frames:
            self.result_frames[self.selected_result_page].configure(border_width=0)
        
        self.selected_result_page = page_num
        
        if page_num in self.result_frames:
            self.result_frames[page_num].configure(border_width=2, border_color="#6f42c1")
        
        self._show_preview(page_num)
    
    def _show_preview(self, page_num: int):
        if page_num < 1 or page_num > len(self.page_images):
            return
        
        img = self.page_images[page_num - 1].copy()
        
        self.preview_canvas.update_idletasks()
        canvas_w = self.preview_canvas.winfo_width() or 250
        canvas_h = self.preview_canvas.winfo_height() or 400
        
        scale = min(canvas_w / img.width, canvas_h / img.height, 1.0)
        new_size = (int(img.width * scale), int(img.height * scale))
        img = img.resize(new_size, Image.Resampling.LANCZOS)
        
        self.preview_tk_image = ImageTk.PhotoImage(img)
        
        self.preview_canvas.delete("all")
        x = (canvas_w - img.width) // 2
        y = (canvas_h - img.height) // 2
        self.preview_canvas.create_image(x, y, anchor="nw", image=self.preview_tk_image)
        
        result = next((r for r in self.results if r["page"] == page_num), None)
        if result:
            tags = result.get("tags", [])
            tags_text = ", ".join(tags) if tags else "No tags"
            self.preview_info.configure(text=f"Page {page_num}: {tags_text}")
    
    def _on_processing_complete(self):
        self.is_running = False
        self.trial_btn.configure(state="normal")
        self.run_btn.configure(state="normal")
        
        tagged_count = sum(1 for r in self.results if r.get("tags"))
        self.progress_label.configure(text=f"Done! {tagged_count}/{len(self.results)} pages tagged")
        
        if self.results:
            self.apply_btn.configure(state="normal")
    
    def _run_trial(self):
        if not self.config.api_key:
            messagebox.showwarning("No API Key", "Set your OpenRouter API key first.\n\nClick the 'API Key' button in the header.")
            return
        
        tag_defs = self._get_tag_definitions()
        if not tag_defs:
            messagebox.showwarning("No Tags", "Add at least one tag definition.")
            return
        
        total_pages = len(self.page_images)
        sample_size = min(20, total_pages)
        page_indices = random.sample(range(total_pages), sample_size)
        page_indices.sort()
        
        self._process_pages(page_indices, tag_defs)
    
    def _run_all(self):
        if not self.config.api_key:
            messagebox.showwarning("No API Key", "Set your OpenRouter API key first.\n\nClick the 'API Key' button in the header.")
            return
        
        tag_defs = self._get_tag_definitions()
        if not tag_defs:
            messagebox.showwarning("No Tags", "Add at least one tag definition.")
            return
        
        if not messagebox.askyesno("Run on All", f"Process all {len(self.page_images)} pages?\n\nThis may take a while and use API credits."):
            return
        
        page_indices = list(range(len(self.page_images)))
        self._process_pages(page_indices, tag_defs)
    
    def _apply_tags(self):
        if not self.results:
            return
        
        count = 0
        for result in self.results:
            if result.get("success") and result.get("tags"):
                page_num = result["page"]
                for tag in result["tags"]:
                    self.source.add_page_tag(page_num, tag)
                    count += 1
        
        self.update_callback()
        self.update_all_callback()
        
        messagebox.showinfo("Applied", f"Applied {count} tags across {len(self.results)} pages.")


class ProjectBrowser(ctk.CTkFrame):
    """Browse and manage projects."""
    
    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self.config = app.config
        
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        
        # Header
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=20, pady=20)
        header.grid_columnconfigure(1, weight=1)
        
        ctk.CTkButton(
            header,
            text="< Back",
            width=80,
            command=self.app.show_welcome
        ).grid(row=0, column=0, padx=(0, 20))
        
        ctk.CTkLabel(
            header,
            text="Projects",
            font=ctk.CTkFont(size=24, weight="bold")
        ).grid(row=0, column=1, sticky="w")
        
        ctk.CTkButton(
            header,
            text="+ New Project",
            width=120,
            fg_color="#28a745",
            hover_color="#218838",
            command=self._new_project
        ).grid(row=0, column=2)
        
        # Projects list
        self.projects_frame = ctk.CTkScrollableFrame(self)
        self.projects_frame.grid(row=1, column=0, sticky="nsew", padx=20, pady=(0, 20))
        self.projects_frame.grid_columnconfigure(0, weight=1)
        
        self._refresh_projects()
    
    def _refresh_projects(self):
        for widget in self.projects_frame.winfo_children():
            widget.destroy()
        
        # Find all projects
        projects = []
        
        # Check projects folder if set
        if self.config.projects_folder and self.config.projects_folder.exists():
            for item in self.config.projects_folder.iterdir():
                if item.is_dir() and (item / "project_meta.json").exists():
                    projects.append(Project(item))
        
        # Also check sources folder for any projects there
        if self.config.sources_folder and self.config.sources_folder.exists():
            self._find_projects_recursive(self.config.sources_folder, projects)
        
        if not projects:
            ctk.CTkLabel(
                self.projects_frame,
                text="No projects yet. Click '+ New Project' to create one.",
                text_color="#888888"
            ).pack(pady=50)
            return
        
        for project in sorted(projects, key=lambda p: p.name.lower()):
            self._create_project_card(project)
    
    def _find_projects_recursive(self, folder: Path, projects: list):
        try:
            for item in folder.iterdir():
                if item.is_dir():
                    if (item / "project_meta.json").exists():
                        projects.append(Project(item))
                    elif not (item / "source_meta.json").exists():
                        self._find_projects_recursive(item, projects)
        except PermissionError:
            pass
    
    def _create_project_card(self, project: Project):
        card = ctk.CTkFrame(self.projects_frame)
        card.pack(fill="x", pady=5)
        card.grid_columnconfigure(1, weight=1)
        
        # Icon
        ctk.CTkLabel(card, text="📋", font=ctk.CTkFont(size=24)).grid(
            row=0, column=0, rowspan=2, padx=15, pady=10
        )
        
        # Name
        ctk.CTkLabel(
            card,
            text=project.name,
            font=ctk.CTkFont(size=16, weight="bold"),
            anchor="w"
        ).grid(row=0, column=1, sticky="w", pady=(10, 0))
        
        # Info
        page_count = len(project.pages)
        ctk.CTkLabel(
            card,
            text=f"{page_count} pages",
            text_color="#888888",
            anchor="w"
        ).grid(row=1, column=1, sticky="w", pady=(0, 10))
        
        # Open button
        ctk.CTkButton(
            card,
            text="Open",
            width=80,
            command=lambda p=project: self.app.show_project_editor(p)
        ).grid(row=0, column=2, rowspan=2, padx=15)
    
    def _new_project(self):
        NewProjectDialog(self, self.app, self._refresh_projects)


class NewProjectDialog(ctk.CTkToplevel):
    """Dialog to create a new project."""
    
    def __init__(self, parent, app, callback):
        super().__init__(parent)
        self.app = app
        self.config = app.config
        self.callback = callback
        
        self.title("New Project")
        self.geometry("500x250")
        self.transient(parent)
        self.grab_set()
        
        self.grid_columnconfigure(0, weight=1)
        
        ctk.CTkLabel(
            self,
            text="Create New Project",
            font=ctk.CTkFont(size=18, weight="bold")
        ).grid(row=0, column=0, pady=20)
        
        # Name
        name_frame = ctk.CTkFrame(self, fg_color="transparent")
        name_frame.grid(row=1, column=0, pady=10, padx=20, sticky="ew")
        name_frame.grid_columnconfigure(1, weight=1)
        
        ctk.CTkLabel(name_frame, text="Name:").grid(row=0, column=0, padx=(0, 10))
        self.name_entry = ctk.CTkEntry(name_frame)
        self.name_entry.grid(row=0, column=1, sticky="ew")
        
        # Location
        loc_frame = ctk.CTkFrame(self, fg_color="transparent")
        loc_frame.grid(row=2, column=0, pady=10, padx=20, sticky="ew")
        loc_frame.grid_columnconfigure(1, weight=1)
        
        ctk.CTkLabel(loc_frame, text="Save in:").grid(row=0, column=0, padx=(0, 10))
        
        default_loc = self.config.projects_folder or self.config.sources_folder or Path.home()
        self.location_path = default_loc
        
        self.location_label = ctk.CTkLabel(loc_frame, text=str(default_loc), text_color="#888888")
        self.location_label.grid(row=0, column=1, sticky="w")
        
        ctk.CTkButton(loc_frame, text="Browse", width=80, command=self._choose_location).grid(row=0, column=2, padx=(10, 0))
        
        # Buttons
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.grid(row=3, column=0, pady=30)
        
        ctk.CTkButton(btn_frame, text="Cancel", width=100, fg_color="transparent", border_width=1, command=self.destroy).pack(side="left", padx=10)
        ctk.CTkButton(btn_frame, text="Create", width=100, fg_color="#28a745", hover_color="#218838", command=self._create).pack(side="left", padx=10)
    
    def _choose_location(self):
        folder = filedialog.askdirectory(title="Choose Project Location", initialdir=str(self.location_path))
        if folder:
            self.location_path = Path(folder)
            self.location_label.configure(text=str(self.location_path))
    
    def _create(self):
        name = self.name_entry.get().strip()
        if not name:
            messagebox.showerror("Error", "Please enter a project name.")
            return
        
        project_path = self.location_path / name
        if project_path.exists():
            messagebox.showerror("Error", f"A project named '{name}' already exists at this location.")
            return
        
        try:
            project_path.mkdir(parents=True)
            
            meta = {
                "name": name,
                "created": datetime.now().isoformat(),
                "pages": []
            }
            save_json(project_path / "project_meta.json", meta)
            
            # Set projects folder if not set
            if not self.config.projects_folder:
                self.config.projects_folder = self.location_path
            
            self.callback()
            self.destroy()
            
            project = Project(project_path)
            self.app.show_project_editor(project)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create project: {e}")


class ProjectEditor(ctk.CTkFrame):
    """Edit a project - add pages from sources, rearrange, export."""
    
    def __init__(self, parent, app, project: Project):
        super().__init__(parent)
        self.app = app
        self.project = project
        self.config = app.config
        
        app.config.add_recent_project(project.path)
        
        self.selected_pages: set[int] = set()
        self.last_clicked_index: Optional[int] = None
        self.page_thumbnails: list[ImageTk.PhotoImage] = []
        self.page_frames: list[ctk.CTkFrame] = []
        self.view_mode = tk.StringVar(value="grid")
        self.pdf_current_page = 0
        
        self.copy_queue: list[int] = []
        self.copy_queue_index = 0
        self.copy_queue_active = False
        
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        
        self._setup_sidebar()
        self._setup_pages_view()
        self._refresh_pages()
        self._bind_keyboard_shortcuts()
    
    def _bind_keyboard_shortcuts(self):
        self.bind("<Delete>", lambda e: self._remove_selected())
        self.bind("<BackSpace>", lambda e: self._remove_selected())
        self.bind("<Up>", lambda e: self._move_up())
        self.bind("<Down>", lambda e: self._move_down())
        self.bind("<Left>", lambda e: self._nav_prev())
        self.bind("<Right>", lambda e: self._nav_next())
        self.bind("<question>", lambda e: self._show_keybind_help())
        self.bind("<Escape>", lambda e: self._on_escape())
        self.bind("<Control-v>", lambda e: self._on_ctrl_v())
        self.bind("<Control-c>", lambda e: self._copy_selected_page())
        self.focus_set()
    
    def _on_escape(self):
        if self.copy_queue_active:
            self._cancel_queue_copy()
        else:
            self._clear_selection()
    
    def _on_ctrl_v(self):
        if self.copy_queue_active:
            self._queue_copy_next()
    
    def _nav_prev(self):
        if self.view_mode.get() == "pdf":
            self._pdf_prev_page()
        else:
            if not self.selected_pages:
                if self.project.pages:
                    self._select_page(0, shift=False)
            else:
                min_idx = min(self.selected_pages)
                if min_idx > 0:
                    self._select_page(min_idx - 1, shift=False)
    
    def _nav_next(self):
        if self.view_mode.get() == "pdf":
            self._pdf_next_page()
        else:
            if not self.selected_pages:
                if self.project.pages:
                    self._select_page(0, shift=False)
            else:
                max_idx = max(self.selected_pages)
                if max_idx < len(self.project.pages) - 1:
                    self._select_page(max_idx + 1, shift=False)
    
    def _clear_selection(self):
        self.selected_pages.clear()
        self.last_clicked_index = None
        self._update_selection_display()
        self._update_selected_label()
    
    def _show_keybind_help(self):
        ProjectKeybindHelpDialog(self)
    
    def _setup_sidebar(self):
        sidebar = ctk.CTkScrollableFrame(self, width=300)
        sidebar.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        
        # Header
        header = ctk.CTkFrame(sidebar, fg_color="transparent")
        header.pack(fill="x", padx=5, pady=10)
        
        ctk.CTkButton(
            header,
            text="< Back",
            width=70,
            command=self.app.show_project_browser
        ).pack(side="left")
        
        ctk.CTkLabel(
            header,
            text=self.project.name,
            font=ctk.CTkFont(size=16, weight="bold"),
            wraplength=180
        ).pack(side="left", padx=10)
        
        # Add from sources
        ctk.CTkLabel(
            sidebar,
            text="Add from Sources",
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(pady=(15, 10), padx=10, anchor="w")
        
        ctk.CTkButton(
            sidebar,
            text="Browse Sources",
            command=self._browse_sources
        ).pack(fill="x", padx=10, pady=3)
        
        ctk.CTkButton(
            sidebar,
            text="Add by Tags",
            command=self._add_by_tags
        ).pack(fill="x", padx=10, pady=3)
        
        # Add custom
        ctk.CTkFrame(sidebar, height=2).pack(fill="x", padx=10, pady=15)
        
        ctk.CTkLabel(
            sidebar,
            text="Add Custom",
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(pady=(5, 10), padx=10, anchor="w")
        
        ctk.CTkButton(
            sidebar,
            text="Add PNG(s) from Files",
            command=self._add_custom_file
        ).pack(fill="x", padx=10, pady=3)
        
        ctk.CTkButton(
            sidebar,
            text="Add from Clipboard",
            command=self._add_from_clipboard
        ).pack(fill="x", padx=10, pady=3)
        
        # Page actions
        ctk.CTkFrame(sidebar, height=2).pack(fill="x", padx=10, pady=15)
        
        ctk.CTkLabel(
            sidebar,
            text="Selected Page",
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(pady=(5, 10), padx=10, anchor="w")
        
        self.selected_label = ctk.CTkLabel(
            sidebar,
            text="No page selected",
            text_color="#888888",
            wraplength=260
        )
        self.selected_label.pack(padx=10, anchor="w")
        
        self.open_source_btn = ctk.CTkButton(
            sidebar,
            text="Open Source Folder",
            width=140,
            height=28,
            fg_color="transparent",
            border_width=1,
            command=self._open_source_folder,
            state="disabled"
        )
        self.open_source_btn.pack(padx=10, pady=(5, 0), anchor="w")
        
        move_frame = ctk.CTkFrame(sidebar, fg_color="transparent")
        move_frame.pack(fill="x", padx=10, pady=5)
        
        ctk.CTkButton(
            move_frame,
            text="Move Up",
            width=90,
            command=self._move_up
        ).pack(side="left", padx=(0, 5))
        
        ctk.CTkButton(
            move_frame,
            text="Move Down",
            width=90,
            command=self._move_down
        ).pack(side="left")
        
        ctk.CTkButton(
            sidebar,
            text="Remove Selected",
            fg_color="#dc3545",
            hover_color="#c82333",
            command=self._remove_selected
        ).pack(fill="x", padx=10, pady=5)
        
        # Copy
        ctk.CTkFrame(sidebar, height=2).pack(fill="x", padx=10, pady=15)
        
        ctk.CTkLabel(
            sidebar,
            text="Copy to Clipboard",
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(pady=(5, 10), padx=10, anchor="w")
        
        ctk.CTkButton(
            sidebar,
            text="Copy Selected Page",
            command=self._copy_selected_page
        ).pack(fill="x", padx=10, pady=3)
        
        ctk.CTkButton(
            sidebar,
            text="Start Queue Copy",
            fg_color="#6f42c1",
            hover_color="#5a32a3",
            command=self._start_queue_copy
        ).pack(fill="x", padx=10, pady=3)
        
        self.queue_frame = ctk.CTkFrame(sidebar, fg_color="#2b2b2b", corner_radius=8)
        self.queue_frame.pack(fill="x", padx=10, pady=5)
        self.queue_frame.pack_forget()
        
        self.queue_label = ctk.CTkLabel(
            self.queue_frame,
            text="Queue: 0/0",
            font=ctk.CTkFont(size=12)
        )
        self.queue_label.pack(pady=(8, 4), padx=10)
        
        self.queue_progress = ctk.CTkProgressBar(self.queue_frame, width=200)
        self.queue_progress.pack(pady=4, padx=10)
        self.queue_progress.set(0)
        
        queue_btns = ctk.CTkFrame(self.queue_frame, fg_color="transparent")
        queue_btns.pack(pady=(4, 8), padx=10)
        
        ctk.CTkButton(
            queue_btns,
            text="Next (Ctrl+V)",
            width=90,
            command=self._queue_copy_next
        ).pack(side="left", padx=2)
        
        ctk.CTkButton(
            queue_btns,
            text="Cancel",
            width=70,
            fg_color="#dc3545",
            hover_color="#c82333",
            command=self._cancel_queue_copy
        ).pack(side="left", padx=2)
        
        # Export
        ctk.CTkFrame(sidebar, height=2).pack(fill="x", padx=10, pady=15)
        
        ctk.CTkLabel(
            sidebar,
            text="Export",
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(pady=(5, 10), padx=10, anchor="w")
        
        ctk.CTkButton(
            sidebar,
            text="Export as PDF",
            fg_color="#28a745",
            hover_color="#218838",
            command=self._export_pdf
        ).pack(fill="x", padx=10, pady=3)
        
        ctk.CTkButton(
            sidebar,
            text="Export as PNG Folder",
            fg_color="#28a745",
            hover_color="#218838",
            command=self._export_pngs
        ).pack(fill="x", padx=10, pady=3)
        
        # Status
        self.status_label = ctk.CTkLabel(sidebar, text="", wraplength=260)
        self.status_label.pack(padx=10, pady=10, anchor="w")
    
    def _setup_pages_view(self):
        self.viewer = ctk.CTkFrame(self)
        self.viewer.grid(row=0, column=1, sticky="nsew", padx=(0, 10), pady=10)
        self.viewer.grid_columnconfigure(0, weight=1)
        self.viewer.grid_rowconfigure(2, weight=1)
        
        info = ctk.CTkFrame(self.viewer, fg_color="transparent")
        info.grid(row=0, column=0, sticky="ew", padx=10, pady=10)
        
        self.page_count_label = ctk.CTkLabel(info, text="0 pages")
        self.page_count_label.pack(side="left")
        
        self.selection_label = ctk.CTkLabel(info, text="", text_color="#28a745")
        self.selection_label.pack(side="left", padx=20)
        
        ctk.CTkButton(
            info,
            text="Clear All",
            width=80,
            fg_color="#dc3545",
            hover_color="#c82333",
            command=self._clear_all
        ).pack(side="right")
        
        view_toggle = ctk.CTkFrame(self.viewer, fg_color="transparent")
        view_toggle.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 5))
        
        ctk.CTkLabel(view_toggle, text="View:").pack(side="left", padx=(0, 10))
        ctk.CTkRadioButton(
            view_toggle, text="Grid", variable=self.view_mode, value="grid",
            command=self._on_view_mode_change
        ).pack(side="left", padx=5)
        ctk.CTkRadioButton(
            view_toggle, text="PDF Preview", variable=self.view_mode, value="pdf",
            command=self._on_view_mode_change
        ).pack(side="left", padx=5)
        
        ctk.CTkLabel(view_toggle, text="  |  Shift+Click for multi-select", text_color="#888888").pack(side="left", padx=20)
        
        ctk.CTkLabel(view_toggle, text="Zoom:", text_color="#888888").pack(side="right", padx=(20, 5))
        ctk.CTkButton(view_toggle, text="-", width=30, command=self._zoom_out).pack(side="right", padx=2)
        ctk.CTkButton(view_toggle, text="+", width=30, command=self._zoom_in).pack(side="right", padx=2)
        
        self.pages_scroll = ctk.CTkScrollableFrame(self.viewer)
        self.pages_scroll.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 10))
        self.viewer.bind("<Configure>", self._on_grid_resize)
        
        self.pdf_view_frame = ctk.CTkFrame(self.viewer)
        
        pdf_nav = ctk.CTkFrame(self.pdf_view_frame, fg_color="transparent")
        pdf_nav.pack(fill="x", pady=10)
        
        self.pdf_prev_btn = ctk.CTkButton(pdf_nav, text="< Prev", width=80, command=self._pdf_prev_page)
        self.pdf_prev_btn.pack(side="left", padx=10)
        
        self.pdf_page_label = ctk.CTkLabel(pdf_nav, text="Page 0 / 0")
        self.pdf_page_label.pack(side="left", expand=True)
        
        self.pdf_next_btn = ctk.CTkButton(pdf_nav, text="Next >", width=80, command=self._pdf_next_page)
        self.pdf_next_btn.pack(side="right", padx=10)
        
        self.pdf_canvas = tk.Canvas(self.pdf_view_frame, bg="#2b2b2b", highlightthickness=0)
        self.pdf_canvas.pack(fill="both", expand=True, padx=10, pady=(0, 10))
    
    def _refresh_pages(self):
        for widget in self.pages_scroll.winfo_children():
            widget.destroy()
        self.page_thumbnails.clear()
        self.page_frames.clear()
        
        pages = self.project.pages
        self.page_count_label.configure(text=f"{len(pages)} pages")
        
        if not pages:
            ctk.CTkLabel(
                self.pages_scroll,
                text="No pages yet. Add pages from sources or custom images.",
                text_color="#888888"
            ).pack(pady=50)
            return
        
        self.thumb_size = getattr(self, 'thumb_size', 120)
        
        self.viewer.update_idletasks()
        container_width = self.viewer.winfo_width() - 40
        if container_width < 100:
            container_width = 700
        
        thumb_total_width = self.thumb_size + 10
        cols = max(1, container_width // thumb_total_width)
        self._current_cols = cols
        
        grid_frame = ctk.CTkFrame(self.pages_scroll, fg_color="transparent")
        grid_frame.pack(anchor="nw", fill="x", expand=True)
        
        for c in range(cols):
            grid_frame.grid_columnconfigure(c, weight=1)
        
        for i, page_info in enumerate(pages):
            row = i // cols
            col = i % cols
            self._create_page_thumbnail(grid_frame, i, page_info, row, col)
        
        self._update_selection_display()
    
    def _create_page_thumbnail(self, parent, index: int, page_info: dict, row: int, col: int):
        size = self.thumb_size
        
        frame = ctk.CTkFrame(parent, border_width=2, border_color="#2b2b2b")
        frame.grid(row=row, column=col, padx=2, pady=4)
        frame.index = index
        self.page_frames.append(frame)
        
        img = self._load_page_image(page_info)
        if img:
            img.thumbnail((size, size))
            tk_img = ImageTk.PhotoImage(img)
            self.page_thumbnails.append(tk_img)
            
            label = ctk.CTkLabel(frame, image=tk_img, text="")
            label.pack(padx=4, pady=(4, 2))
        else:
            placeholder = ctk.CTkLabel(frame, text="[Error]", width=size, height=size, text_color="#888888")
            placeholder.pack(padx=4, pady=(4, 2))
        
        source_name = page_info.get("source_name", "Custom")
        page_num = page_info.get("page", "")
        text = f"#{index + 1}"
        if page_num:
            text += f" p{page_num}"
        
        ctk.CTkLabel(
            frame,
            text=text,
            font=ctk.CTkFont(size=10),
            wraplength=size
        ).pack(pady=(0, 4))
        
        frame.bind("<Button-1>", lambda e, idx=index: self._on_thumb_click(e, idx))
        frame.bind("<Shift-Button-1>", lambda e, idx=index: self._on_thumb_shift_click(e, idx))
        frame.bind("<ButtonPress-1>", lambda e, idx=index: self._on_drag_start(e, idx))
        frame.bind("<B1-Motion>", self._on_drag_motion)
        frame.bind("<ButtonRelease-1>", self._on_drag_end)
        
        for child in frame.winfo_children():
            child.bind("<Button-1>", lambda e, idx=index: self._on_thumb_click(e, idx))
            child.bind("<Shift-Button-1>", lambda e, idx=index: self._on_thumb_shift_click(e, idx))
    
    def _load_page_image(self, page_info: dict) -> Optional[Image.Image]:
        try:
            if page_info.get("type") == "custom":
                # Custom PNG
                path = self.project.path / page_info.get("filename", "")
                if path.exists():
                    return Image.open(path)
            else:
                # From source
                source_path = Path(page_info.get("source", ""))
                if not source_path.exists():
                    return None
                
                source = Source(source_path)
                page_num = page_info.get("page", 1)
                
                if source.source_type == "pdf":
                    pdf_path = source.pdf_path
                    if pdf_path and pdf_path.exists():
                        doc = fitz.open(str(pdf_path))
                        if page_num <= len(doc):
                            page = doc[page_num - 1]
                            pix = page.get_pixmap(matrix=fitz.Matrix(0.5, 0.5))
                            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                            
                            # Apply crop
                            crop = source.get_page_crop(page_num)
                            w, h = img.size
                            scale = 150 / 72 * 0.5  # Match the low-res render
                            left = int(crop["left"] * scale / (150/72))
                            right = int(crop["right"] * scale / (150/72))
                            top = int(crop["top"] * scale / (150/72))
                            bottom = int(crop["bottom"] * scale / (150/72))
                            img = img.crop((left, top, w - right, h - bottom))
                            
                            doc.close()
                            return img
                else:
                    # PNG folder
                    pngs = sorted(list(source.path.glob("*.png")) + list(source.path.glob("*.PNG")))
                    if page_num <= len(pngs):
                        return Image.open(pngs[page_num - 1])
        except Exception as e:
            print(f"Error loading page image: {e}")
        
        return None
    
    def _select_page(self, index: int, shift: bool = False):
        if shift and self.last_clicked_index is not None:
            start = min(self.last_clicked_index, index)
            end = max(self.last_clicked_index, index)
            self.selected_pages = set(range(start, end + 1))
        else:
            self.selected_pages = {index}
            self.last_clicked_index = index
        
        self._update_selection_display()
        self._update_selected_label()
        
        if self.view_mode.get() == "pdf":
            self.pdf_current_page = index
            self._update_pdf_view()
    
    def _on_thumb_click(self, event, index: int):
        self._select_page(index, shift=False)
    
    def _on_thumb_shift_click(self, event, index: int):
        self._select_page(index, shift=True)
    
    def _update_selection_display(self):
        for frame in self.page_frames:
            if hasattr(frame, 'index') and frame.index in self.selected_pages:
                frame.configure(border_color="#28a745")
            else:
                frame.configure(border_color="#2b2b2b")
        
        count = len(self.selected_pages)
        if count > 1:
            self.selection_label.configure(text=f"{count} pages selected")
        elif count == 1:
            self.selection_label.configure(text="1 page selected")
        else:
            self.selection_label.configure(text="")
    
    def _update_selected_label(self):
        if not self.selected_pages:
            self.selected_label.configure(text="No page selected")
            self.open_source_btn.configure(state="disabled")
            self._current_source_path = None
            return
        
        if len(self.selected_pages) == 1:
            index = list(self.selected_pages)[0]
            page = self.project.pages[index]
            source_name = page.get("source_name", "Custom")
            page_num = page.get("page", "")
            source_path = page.get("source", "")
            
            self._current_source_path = source_path if source_path else None
            
            if source_path and self.config.sources_folder:
                try:
                    rel_path = Path(source_path).relative_to(self.config.sources_folder)
                    path_str = str(rel_path)
                except ValueError:
                    path_str = source_path
            else:
                path_str = source_name
            
            text = f"#{index + 1}: {path_str}"
            if page_num:
                text += f" (page {page_num})"
            self.selected_label.configure(text=text)
            
            if source_path and Path(source_path).exists():
                self.open_source_btn.configure(state="normal")
            else:
                self.open_source_btn.configure(state="disabled")
        else:
            indices = sorted(self.selected_pages)
            self.selected_label.configure(text=f"Selected: {', '.join(f'#{i+1}' for i in indices[:5])}{'...' if len(indices) > 5 else ''}")
            self.open_source_btn.configure(state="disabled")
            self._current_source_path = None
    
    def _open_source_folder(self):
        if not self._current_source_path:
            return
        
        source_path = Path(self._current_source_path)
        if source_path.exists():
            if os.name == 'nt':
                os.startfile(source_path)
            elif os.uname().sysname == 'Darwin':
                subprocess.run(['open', str(source_path)])
            else:
                subprocess.run(['xdg-open', str(source_path)])
    
    def _on_drag_start(self, event, index: int):
        self.drag_start_index = index
        self.drag_data = {"x": event.x, "y": event.y, "dragging": False}
    
    def _on_drag_motion(self, event):
        if not hasattr(self, 'drag_data'):
            return
        dx = abs(event.x - self.drag_data["x"])
        dy = abs(event.y - self.drag_data["y"])
        if dx > 10 or dy > 10:
            self.drag_data["dragging"] = True
            self.configure(cursor="fleur")
    
    def _on_drag_end(self, event):
        if not hasattr(self, 'drag_data') or not self.drag_data.get("dragging"):
            self.configure(cursor="")
            return
        
        self.configure(cursor="")
        
        target_index = self._get_drop_target(event)
        if target_index is not None and target_index != self.drag_start_index:
            self.project.move_page(self.drag_start_index, target_index)
            self.selected_pages = {target_index}
            self.last_clicked_index = target_index
            self._refresh_pages()
        
        self.drag_data = None
    
    def _get_drop_target(self, event) -> Optional[int]:
        widget = event.widget.winfo_containing(event.x_root, event.y_root)
        while widget:
            if hasattr(widget, 'index'):
                return widget.index
            widget = widget.master
        return None
    
    def _on_view_mode_change(self):
        if self.view_mode.get() == "pdf":
            self.pages_scroll.grid_forget()
            self.pdf_view_frame.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 10))
            if self.selected_pages:
                self.pdf_current_page = min(self.selected_pages)
            else:
                self.pdf_current_page = 0
            self._update_pdf_view()
        else:
            self.pdf_view_frame.grid_forget()
            self.pages_scroll.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 10))
    
    def _zoom_in(self):
        self.thumb_size = min(300, getattr(self, 'thumb_size', 120) + 30)
        self._refresh_pages()
    
    def _zoom_out(self):
        self.thumb_size = max(60, getattr(self, 'thumb_size', 120) - 30)
        self._refresh_pages()
    
    def _on_grid_resize(self, event):
        if self.view_mode.get() != "grid":
            return
        if not self.project.pages:
            return
            
        container_width = event.width - 40
        if container_width < 100:
            return
            
        thumb_total = self.thumb_size + 10
        new_cols = max(1, container_width // thumb_total)
        old_cols = getattr(self, '_current_cols', 0)
        
        if new_cols != old_cols:
            self._refresh_pages()
    
    def _update_pdf_view(self):
        if not self.project.pages:
            self.pdf_page_label.configure(text="No pages")
            return
        
        self.pdf_current_page = max(0, min(self.pdf_current_page, len(self.project.pages) - 1))
        page_info = self.project.pages[self.pdf_current_page]
        
        self.pdf_page_label.configure(text=f"Page {self.pdf_current_page + 1} / {len(self.project.pages)}")
        
        img = self._load_full_page_image(page_info)
        if img:
            canvas_width = self.pdf_canvas.winfo_width() or 600
            canvas_height = self.pdf_canvas.winfo_height() or 400
            
            scale = min(canvas_width / img.width, canvas_height / img.height, 1.0)
            new_size = (int(img.width * scale), int(img.height * scale))
            img = img.resize(new_size, Image.Resampling.LANCZOS)
            
            self.pdf_tk_image = ImageTk.PhotoImage(img)
            self.pdf_canvas.delete("all")
            x = (canvas_width - img.width) // 2
            y = (canvas_height - img.height) // 2
            self.pdf_canvas.create_image(x, y, anchor="nw", image=self.pdf_tk_image)
    
    def _pdf_prev_page(self):
        if self.pdf_current_page > 0:
            self.pdf_current_page -= 1
            self.selected_pages = {self.pdf_current_page}
            self._update_pdf_view()
            self._update_selected_label()
    
    def _pdf_next_page(self):
        if self.pdf_current_page < len(self.project.pages) - 1:
            self.pdf_current_page += 1
            self.selected_pages = {self.pdf_current_page}
            self._update_pdf_view()
            self._update_selected_label()
    
    def _move_up(self):
        if not self.selected_pages:
            return
        
        indices = sorted(self.selected_pages)
        if indices[0] <= 0:
            return
        
        for idx in indices:
            self.project.move_page(idx, idx - 1)
        
        self.selected_pages = {i - 1 for i in self.selected_pages}
        self._refresh_pages()
    
    def _move_down(self):
        if not self.selected_pages:
            return
        
        indices = sorted(self.selected_pages, reverse=True)
        if indices[0] >= len(self.project.pages) - 1:
            return
        
        for idx in indices:
            self.project.move_page(idx, idx + 1)
        
        self.selected_pages = {i + 1 for i in self.selected_pages}
        self._refresh_pages()
    
    def _remove_selected(self):
        if not self.selected_pages:
            return
        
        for idx in sorted(self.selected_pages, reverse=True):
            self.project.remove_page(idx)
        
        self.selected_pages.clear()
        self.last_clicked_index = None
        self._update_selected_label()
        self._refresh_pages()
    
    def _copy_selected_page(self):
        if not self.selected_pages:
            messagebox.showinfo("No Selection", "Select a page first.")
            return
        
        idx = min(self.selected_pages)
        try:
            img = self._load_page_image(idx)
            if img:
                copy_image_to_clipboard(img)
                self.status_label.configure(text=f"Copied page {idx + 1} to clipboard")
        except Exception as e:
            messagebox.showerror("Error", f"Copy failed: {e}")
    
    def _start_queue_copy(self):
        if not self.selected_pages:
            messagebox.showinfo("No Selection", "Select pages to queue for copying.\nUse Shift+Click for range selection.")
            return
        
        self.copy_queue = sorted(self.selected_pages)
        self.copy_queue_index = 0
        self.copy_queue_active = True
        
        self._update_queue_ui()
        self.queue_frame.pack(fill="x", padx=10, pady=5)
        
        self._queue_copy_next()
    
    def _queue_copy_next(self):
        if not self.copy_queue_active or self.copy_queue_index >= len(self.copy_queue):
            self._cancel_queue_copy()
            self.status_label.configure(text="Queue complete!")
            return
        
        idx = self.copy_queue[self.copy_queue_index]
        try:
            img = self._load_page_image(idx)
            if img:
                copy_image_to_clipboard(img)
                self.status_label.configure(text=f"Copied {self.copy_queue_index + 1}/{len(self.copy_queue)} (page {idx + 1})")
                self.copy_queue_index += 1
                self._update_queue_ui()
        except Exception as e:
            messagebox.showerror("Error", f"Copy failed: {e}")
    
    def _cancel_queue_copy(self):
        self.copy_queue_active = False
        self.copy_queue = []
        self.copy_queue_index = 0
        self.queue_frame.pack_forget()
    
    def _update_queue_ui(self):
        total = len(self.copy_queue)
        current = self.copy_queue_index
        self.queue_label.configure(text=f"Queue: {current}/{total}")
        self.queue_progress.set(current / total if total > 0 else 0)
    
    def _load_page_image(self, idx: int) -> Optional[Image.Image]:
        if idx < 0 or idx >= len(self.project.pages):
            return None
        
        page = self.project.pages[idx]
        
        if page.get("type") == "custom":
            filepath = self.project.path / page["filename"]
            if filepath.exists():
                return Image.open(filepath)
        else:
            source_path = Path(page.get("source_path", ""))
            if source_path.exists():
                source = Source(source_path)
                page_num = page.get("page", 1)
                
                source_type = source.meta.get("source_type", "pdf")
                if source_type == "pdf":
                    pdf_files = sorted(source_path.glob("*.pdf"))
                    if pdf_files:
                        doc = fitz.open(pdf_files[0])
                        if page_num <= len(doc):
                            pdf_page = doc[page_num - 1]
                            mat = fitz.Matrix(2.0, 2.0)
                            pix = pdf_page.get_pixmap(matrix=mat)
                            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                            
                            crop = source.get_page_crop(page_num)
                            w, h = img.size
                            left = int(crop["left"])
                            right = int(crop["right"])
                            top = int(crop["top"])
                            bottom = int(crop["bottom"])
                            img = img.crop((left, top, w - right, h - bottom))
                            
                            doc.close()
                            return img
                else:
                    images_folder = source_path / "images"
                    if images_folder.exists():
                        img_files = sorted([f for f in images_folder.iterdir() if f.suffix.lower() in ['.png', '.jpg', '.jpeg']])
                        if page_num <= len(img_files):
                            img = Image.open(img_files[page_num - 1])
                            
                            crop = source.get_page_crop(page_num)
                            w, h = img.size
                            left = int(crop["left"])
                            right = int(crop["right"])
                            top = int(crop["top"])
                            bottom = int(crop["bottom"])
                            return img.crop((left, top, w - right, h - bottom))
        
        return None
    
    def _clear_all(self):
        if not self.project.pages:
            return
        
        if messagebox.askyesno("Clear All", "Remove all pages from this project?"):
            self.project.clear_pages()
            self.selected_pages.clear()
            self.last_clicked_index = None
            self._update_selected_label()
            self._refresh_pages()
    
    def _browse_sources(self):
        SourcePickerDialog(self, self.app, self.project, self._refresh_pages)
    
    def _add_by_tags(self):
        AddByTagsDialog(self, self.app, self.project, self._refresh_pages)
    
    def _add_custom_file(self):
        filepaths = filedialog.askopenfilenames(
            title="Select PNG(s)",
            filetypes=[("PNG files", "*.png"), ("Image files", "*.png *.jpg *.jpeg"), ("All files", "*.*")]
        )
        if not filepaths:
            return
        
        try:
            added = 0
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            for filepath in filepaths:
                src = Path(filepath)
                dest = self.project.path / f"custom_{timestamp}_{added}_{src.name}"
                shutil.copy2(src, dest)
                
                self.project.add_page({
                    "type": "custom",
                    "filename": dest.name,
                    "source_name": src.stem
                })
                added += 1
            
            self._refresh_pages()
            self.status_label.configure(text=f"Added {added} image(s)")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to add image: {e}")
    
    def _add_from_clipboard(self):
        try:
            # Try to get image from clipboard using PIL
            from PIL import ImageGrab
            img = ImageGrab.grabclipboard()
            
            if img is None:
                messagebox.showinfo("No Image", "No image found in clipboard.")
                return
            
            # Save to project folder
            filename = f"clipboard_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
            dest = self.project.path / filename
            img.save(dest, "PNG")
            
            self.project.add_page({
                "type": "custom",
                "filename": filename,
                "source_name": "Clipboard"
            })
            
            self._refresh_pages()
            self.status_label.configure(text="Added image from clipboard")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to get clipboard image: {e}")
    
    def _export_pdf(self):
        if not self.project.pages:
            messagebox.showinfo("No Pages", "Add some pages to the project first.")
            return
        
        images = []
        for page_info in self.project.pages:
            img = self._load_full_page_image(page_info)
            if img:
                images.append(img)
        
        if not images:
            messagebox.showerror("Error", "Could not load any pages.")
            return
        
        ExportPDFDialog(self, images, self.project.name)
    
    def _export_pngs(self):
        if not self.project.pages:
            messagebox.showinfo("No Pages", "Add some pages to the project first.")
            return
        
        folder = filedialog.askdirectory(title="Select Output Folder")
        if not folder:
            return
        
        try:
            output_path = Path(folder)
            count = 0
            
            for i, page_info in enumerate(self.project.pages):
                img = self._load_full_page_image(page_info)
                if img:
                    output_file = output_path / f"page_{i + 1:03d}.png"
                    img.save(output_file, "PNG")
                    count += 1
            
            self.status_label.configure(text=f"Exported {count} PNGs")
            messagebox.showinfo("Export Complete", f"Exported {count} pages to:\n{folder}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Export failed: {e}")
    
    def _load_full_page_image(self, page_info: dict) -> Optional[Image.Image]:
        """Load full-resolution page image with crop applied."""
        try:
            if page_info.get("type") == "custom":
                path = self.project.path / page_info.get("filename", "")
                if path.exists():
                    return Image.open(path)
            else:
                source_path = Path(page_info.get("source", ""))
                if not source_path.exists():
                    return None
                
                source = Source(source_path)
                page_num = page_info.get("page", 1)
                
                if source.source_type == "pdf":
                    pdf_path = source.pdf_path
                    if pdf_path and pdf_path.exists():
                        doc = fitz.open(str(pdf_path))
                        if page_num <= len(doc):
                            page = doc[page_num - 1]
                            zoom = 150 / 72
                            pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
                            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                            
                            # Apply crop
                            crop = source.get_page_crop(page_num)
                            w, h = img.size
                            left = int(crop["left"])
                            right = int(crop["right"])
                            top = int(crop["top"])
                            bottom = int(crop["bottom"])
                            img = img.crop((left, top, w - right, h - bottom))
                            
                            doc.close()
                            return img
                else:
                    pngs = sorted(list(source.path.glob("*.png")) + list(source.path.glob("*.PNG")))
                    pngs = [p for p in pngs if p.name != "source_meta.json"]
                    if page_num <= len(pngs):
                        img = Image.open(pngs[page_num - 1])
                        
                        # Apply crop
                        crop = source.get_page_crop(page_num)
                        w, h = img.size
                        left = int(crop["left"])
                        right = int(crop["right"])
                        top = int(crop["top"])
                        bottom = int(crop["bottom"])
                        img = img.crop((left, top, w - right, h - bottom))
                        
                        return img
        except Exception as e:
            print(f"Error loading full page image: {e}")
        
        return None


class SourcePickerDialog(ctk.CTkToplevel):
    """Dialog to pick pages from sources."""
    
    def __init__(self, parent, app, project: Project, callback):
        super().__init__(parent)
        self.app = app
        self.project = project
        self.callback = callback
        self.config = app.config
        
        self.title("Add from Sources")
        self.geometry("800x600")
        self.transient(parent)
        self.grab_set()
        
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(1, weight=1)
        
        # Header
        ctk.CTkLabel(
            self,
            text="Select Pages from Sources",
            font=ctk.CTkFont(size=18, weight="bold")
        ).grid(row=0, column=0, columnspan=2, pady=20)
        
        # Sources list
        sources_frame = ctk.CTkFrame(self)
        sources_frame.grid(row=1, column=0, sticky="nsew", padx=(20, 10), pady=(0, 20))
        sources_frame.grid_columnconfigure(0, weight=1)
        sources_frame.grid_rowconfigure(1, weight=1)
        
        ctk.CTkLabel(sources_frame, text="Sources", font=ctk.CTkFont(weight="bold")).grid(
            row=0, column=0, sticky="w", padx=10, pady=10
        )
        
        self.sources_list = ctk.CTkScrollableFrame(sources_frame)
        self.sources_list.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        
        # Pages view
        pages_frame = ctk.CTkFrame(self)
        pages_frame.grid(row=1, column=1, sticky="nsew", padx=(10, 20), pady=(0, 20))
        pages_frame.grid_columnconfigure(0, weight=1)
        pages_frame.grid_rowconfigure(1, weight=1)
        
        ctk.CTkLabel(pages_frame, text="Pages", font=ctk.CTkFont(weight="bold")).grid(
            row=0, column=0, sticky="w", padx=10, pady=10
        )
        
        self.pages_list = ctk.CTkScrollableFrame(pages_frame)
        self.pages_list.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        
        # Buttons
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.grid(row=2, column=0, columnspan=2, pady=10)
        
        ctk.CTkButton(btn_frame, text="Close", command=self.destroy).pack(side="right", padx=10)
        
        self.selected_source: Optional[Source] = None
        self.selected_pages: set[int] = set()
        self.last_clicked_page: Optional[int] = None
        
        self.existing_pages = {(p.get("source"), p.get("page")) for p in self.project.pages}
        
        self._load_sources()
    
    def _load_sources(self):
        sources_folder = self.config.sources_folder
        if not sources_folder or not sources_folder.exists():
            return
        
        sources = []
        self._find_sources_recursive(sources_folder, sources)
        
        for source in sorted(sources, key=lambda s: s.name.lower()):
            btn = ctk.CTkButton(
                self.sources_list,
                text=source.name,
                anchor="w",
                fg_color="transparent",
                text_color=("#000", "#fff"),
                hover_color=("#e0e0e0", "#3a3a3a"),
                command=lambda s=source: self._select_source(s)
            )
            btn.pack(fill="x", pady=2)
    
    def _find_sources_recursive(self, folder: Path, sources: list):
        try:
            for item in folder.iterdir():
                if item.is_dir():
                    if (item / "source_meta.json").exists():
                        sources.append(Source(item))
                    else:
                        self._find_sources_recursive(item, sources)
        except PermissionError:
            pass
    
    def _select_source(self, source: Source):
        self.selected_source = source
        self.selected_pages.clear()
        self.last_clicked_page = None
        self.page_checkboxes = {}
        
        for widget in self.pages_list.winfo_children():
            widget.destroy()
        
        start, end = source.get_page_range()
        total = source.get_page_count()
        
        for page_num in range(start, min(end + 1, total + 1)):
            tags = source.get_page_tags(page_num)
            tag_str = f" [{', '.join(tags)}]" if tags else ""
            override = " *" if source.has_page_crop_override(page_num) else ""
            
            is_dupe = (str(source.path), page_num) in self.existing_pages
            dupe_str = " [IN PROJECT]" if is_dupe else ""
            
            var = tk.BooleanVar()
            text_color = "#ff9800" if is_dupe else ("#000", "#fff")
            
            row = ctk.CTkFrame(self.pages_list, fg_color="transparent")
            row.pack(fill="x", pady=1)
            
            cb = ctk.CTkCheckBox(
                row,
                text=f"Page {page_num}{tag_str}{override}{dupe_str}",
                variable=var,
                text_color=text_color,
                command=lambda pn=page_num, v=var: self._toggle_page(pn, v)
            )
            cb.pack(side="left", anchor="w")
            cb.bind("<Shift-Button-1>", lambda e, pn=page_num: self._shift_select_page(pn))
            
            self.page_checkboxes[page_num] = (var, cb)
        
        btn_frame = ctk.CTkFrame(self.pages_list, fg_color="transparent")
        btn_frame.pack(fill="x", pady=10)
        
        ctk.CTkButton(
            btn_frame,
            text="Select All",
            width=100,
            command=self._select_all_pages
        ).pack(side="left", padx=5)
        
        ctk.CTkButton(
            btn_frame,
            text="Select None",
            width=100,
            command=self._select_no_pages
        ).pack(side="left", padx=5)
        
        ctk.CTkButton(
            btn_frame,
            text="Add Selected",
            fg_color="#28a745",
            hover_color="#218838",
            command=self._add_selected
        ).pack(side="right", padx=5)
    
    def _shift_select_page(self, page_num: int):
        if self.last_clicked_page is not None:
            start = min(self.last_clicked_page, page_num)
            end = max(self.last_clicked_page, page_num)
            for pn in range(start, end + 1):
                if pn in self.page_checkboxes:
                    var, cb = self.page_checkboxes[pn]
                    var.set(True)
                    self.selected_pages.add(pn)
        self.last_clicked_page = page_num
    
    def _select_all_pages(self):
        for pn, (var, cb) in self.page_checkboxes.items():
            var.set(True)
            self.selected_pages.add(pn)
    
    def _select_no_pages(self):
        for pn, (var, cb) in self.page_checkboxes.items():
            var.set(False)
        self.selected_pages.clear()
    
    def _toggle_page(self, page_num: int, var: tk.BooleanVar):
        if var.get():
            self.selected_pages.add(page_num)
        else:
            self.selected_pages.discard(page_num)
        self.last_clicked_page = page_num
    
    def _add_selected(self):
        if not self.selected_source or not self.selected_pages:
            return
        
        dupes = []
        new_pages = []
        for page_num in sorted(self.selected_pages):
            key = (str(self.selected_source.path), page_num)
            if key in self.existing_pages:
                dupes.append(page_num)
            new_pages.append({
                "type": "source",
                "source": str(self.selected_source.path),
                "source_name": self.selected_source.name,
                "page": page_num
            })
        
        if dupes:
            dupe_str = ", ".join(str(p) for p in dupes[:5])
            if len(dupes) > 5:
                dupe_str += f"... ({len(dupes)} total)"
            if not messagebox.askyesno(
                "Duplicate Pages",
                f"Pages {dupe_str} are already in the project.\n\nAdd them anyway?"
            ):
                return
        
        self.project.add_pages(new_pages)
        self.callback()
        
        for p in new_pages:
            self.existing_pages.add((p["source"], p["page"]))
        
        messagebox.showinfo("Added", f"Added {len(new_pages)} pages to project.")
        self.selected_pages.clear()
        
        self._select_source(self.selected_source)


class AddByTagsDialog(ctk.CTkToplevel):
    """Dialog to add pages by tags with filters."""
    
    def __init__(self, parent, app, project: Project, callback):
        super().__init__(parent)
        self.app = app
        self.project = project
        self.callback = callback
        self.config = app.config
        
        self.title("Add by Tags")
        self.geometry("600x650")
        self.transient(parent)
        self.grab_set()
        
        self.grid_columnconfigure(0, weight=1)
        
        ctk.CTkLabel(
            self,
            text="Add Pages by Tags",
            font=ctk.CTkFont(size=18, weight="bold")
        ).grid(row=0, column=0, pady=20)
        
        self.all_tags = self._collect_all_tags()
        
        tag_frame = ctk.CTkFrame(self, fg_color="transparent")
        tag_frame.grid(row=1, column=0, pady=10, padx=20, sticky="ew")
        
        ctk.CTkLabel(tag_frame, text="Selected tags:").pack(anchor="w")
        self.tags_entry = ctk.CTkEntry(tag_frame, width=500, placeholder_text="Click tags below or type comma-separated")
        self.tags_entry.pack(fill="x", pady=5)
        self.tags_entry.bind("<KeyRelease>", self._on_tags_changed)
        
        self.warning_label = ctk.CTkLabel(tag_frame, text="", text_color="#ff9800", font=ctk.CTkFont(size=11))
        self.warning_label.pack(anchor="w")
        
        if self.all_tags:
            ctk.CTkLabel(tag_frame, text="Click to add:", text_color="#888888").pack(anchor="w", pady=(10, 5))
            tags_buttons_frame = ctk.CTkFrame(tag_frame, fg_color="transparent")
            tags_buttons_frame.pack(anchor="w", fill="x")
            
            for tag in sorted(self.all_tags):
                btn = ctk.CTkButton(
                    tags_buttons_frame,
                    text=tag,
                    width=len(tag) * 10 + 20,
                    height=28,
                    fg_color="transparent",
                    border_width=1,
                    command=lambda t=tag: self._add_tag(t)
                )
                btn.pack(side="left", padx=2, pady=2)
        
        mode_frame = ctk.CTkFrame(self, fg_color="transparent")
        mode_frame.grid(row=2, column=0, pady=10, padx=20, sticky="ew")
        
        ctk.CTkLabel(mode_frame, text="Match mode:").pack(side="left", padx=(0, 10))
        
        self.mode_var = tk.StringVar(value="any")
        ctk.CTkRadioButton(mode_frame, text="Any tag (OR)", variable=self.mode_var, value="any").pack(side="left", padx=10)
        ctk.CTkRadioButton(mode_frame, text="All tags (AND)", variable=self.mode_var, value="all").pack(side="left", padx=10)
        
        filter_frame = ctk.CTkFrame(self, fg_color="transparent")
        filter_frame.grid(row=3, column=0, pady=10, padx=20, sticky="ew")
        
        ctk.CTkLabel(filter_frame, text="Options:", font=ctk.CTkFont(weight="bold")).pack(anchor="w", pady=(0, 5))
        
        limit_row = ctk.CTkFrame(filter_frame, fg_color="transparent")
        limit_row.pack(fill="x", pady=2)
        
        self.limit_var = tk.BooleanVar(value=False)
        ctk.CTkCheckBox(limit_row, text="Limit to", variable=self.limit_var, width=80).pack(side="left")
        self.limit_entry = ctk.CTkEntry(limit_row, width=60)
        self.limit_entry.pack(side="left", padx=5)
        self.limit_entry.insert(0, "10")
        ctk.CTkLabel(limit_row, text="pages", text_color="#888888").pack(side="left", padx=5)
        
        self.random_var = tk.BooleanVar(value=False)
        ctk.CTkCheckBox(filter_frame, text="Randomize order", variable=self.random_var).pack(anchor="w", pady=2)
        
        self.skip_dupes_var = tk.BooleanVar(value=True)
        ctk.CTkCheckBox(filter_frame, text="Skip pages already in project", variable=self.skip_dupes_var).pack(anchor="w", pady=2)
        
        self.preview_label = ctk.CTkLabel(self, text="", text_color="#888888")
        self.preview_label.grid(row=4, column=0, pady=10)
        
        ctk.CTkButton(self, text="Preview", command=self._preview).grid(row=5, column=0, pady=5)
        
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.grid(row=6, column=0, pady=20)
        
        ctk.CTkButton(btn_frame, text="Cancel", fg_color="transparent", border_width=1, command=self.destroy).pack(side="left", padx=10)
        ctk.CTkButton(btn_frame, text="Add Pages", fg_color="#28a745", hover_color="#218838", command=self._add).pack(side="left", padx=10)
        
        self.matching_pages: list[dict] = []
    
    def _add_tag(self, tag: str):
        current = self.tags_entry.get().strip()
        current_tags = {t.strip().lower() for t in current.split(",") if t.strip()}
        
        if tag.lower() not in current_tags:
            if current:
                self.tags_entry.insert("end", f", {tag}")
            else:
                self.tags_entry.insert(0, tag)
        self._on_tags_changed(None)
    
    def _on_tags_changed(self, event):
        current = self.tags_entry.get().strip()
        if not current:
            self.warning_label.configure(text="")
            return
        
        typed_tags = {t.strip().lower() for t in current.split(",") if t.strip()}
        available_lower = {t.lower() for t in self.all_tags}
        
        unknown = typed_tags - available_lower
        if unknown:
            self.warning_label.configure(text=f"Unknown: {', '.join(sorted(unknown))}")
        else:
            self.warning_label.configure(text="")
    
    def _collect_all_tags(self) -> set[str]:
        tags = set()
        sources_folder = self.config.sources_folder
        if sources_folder and sources_folder.exists():
            self._collect_tags_recursive(sources_folder, tags)
        return tags
    
    def _collect_tags_recursive(self, folder: Path, tags: set):
        try:
            for item in folder.iterdir():
                if item.is_dir():
                    if (item / "source_meta.json").exists():
                        source = Source(item)
                        tags.update(source.get_all_tags())
                    else:
                        self._collect_tags_recursive(item, tags)
        except PermissionError:
            pass
    
    def _find_matching_pages(self) -> list[dict]:
        tags_input = self.tags_entry.get().strip()
        if not tags_input:
            return []
        
        search_tags = {t.strip().lower() for t in tags_input.split(",") if t.strip()}
        mode = self.mode_var.get()
        
        matches = []
        sources_folder = self.config.sources_folder
        if sources_folder and sources_folder.exists():
            self._find_matches_recursive(sources_folder, search_tags, mode, matches)
        
        return matches
    
    def _find_matches_recursive(self, folder: Path, search_tags: set, mode: str, matches: list):
        try:
            for item in folder.iterdir():
                if item.is_dir():
                    if (item / "source_meta.json").exists():
                        source = Source(item)
                        start, end = source.get_page_range()
                        total = source.get_page_count()
                        
                        for page_num in range(start, min(end + 1, total + 1)):
                            page_tags = {t.lower() for t in source.get_page_tags(page_num)}
                            
                            if mode == "any":
                                if search_tags & page_tags:
                                    matches.append({
                                        "type": "source",
                                        "source": str(source.path),
                                        "source_name": source.name,
                                        "page": page_num
                                    })
                            else:  # all
                                if search_tags <= page_tags:
                                    matches.append({
                                        "type": "source",
                                        "source": str(source.path),
                                        "source_name": source.name,
                                        "page": page_num
                                    })
                    else:
                        self._find_matches_recursive(item, search_tags, mode, matches)
        except PermissionError:
            pass
    
    def _apply_filters(self, pages: list[dict]) -> list[dict]:
        result = pages.copy()
        
        if self.skip_dupes_var.get():
            existing = {(p.get("source"), p.get("page")) for p in self.project.pages}
            result = [p for p in result if (p.get("source"), p.get("page")) not in existing]
        
        if self.random_var.get():
            random.shuffle(result)
        
        if self.limit_var.get():
            try:
                n = int(self.limit_entry.get())
                result = result[:n]
            except ValueError:
                pass
        
        return result
    
    def _preview(self):
        matches = self._find_matching_pages()
        filtered = self._apply_filters(matches)
        
        self.matching_pages = filtered
        
        skipped = len(matches) - len(filtered)
        if skipped > 0 and self.skip_dupes_var.get():
            self.preview_label.configure(text=f"Found {len(matches)} matching, {skipped} duplicates skipped, will add {len(filtered)}")
        else:
            self.preview_label.configure(text=f"Found {len(matches)} matching pages, will add {len(filtered)}")
    
    def _add(self):
        if not self.matching_pages:
            self._preview()
        
        if not self.matching_pages:
            messagebox.showinfo("No Matches", "No pages match the specified tags.")
            return
        
        self.project.add_pages(self.matching_pages)
        self.callback()
        
        messagebox.showinfo("Added", f"Added {len(self.matching_pages)} pages to project.")
        self.destroy()


class ExportPDFDialog(ctk.CTkToplevel):
    PAGE_SIZES = {
        "A4": (595, 842),
        "A4 Landscape": (842, 595),
        "Letter": (612, 792),
        "Letter Landscape": (792, 612),
        "A3": (842, 1191),
        "A3 Landscape": (1191, 842),
        "A5": (420, 595),
        "A5 Landscape": (595, 420),
        "Original Size": None,
    }
    
    def __init__(self, parent, images: list[Image.Image], project_name: str):
        super().__init__(parent)
        self.images = images
        self.project_name = project_name
        self.current_page = 0
        
        self.title("Export PDF - Preview")
        self.geometry("900x700")
        self.transient(parent)
        self.grab_set()
        
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        
        self.page_size_var = tk.StringVar(value="A4")
        self.fit_mode_var = tk.StringVar(value="fit")
        self.margin_var = tk.IntVar(value=20)
        self.quality_var = tk.IntVar(value=90)
        
        self._setup_sidebar()
        self._setup_preview()
        self._update_preview()
    
    def _setup_sidebar(self):
        sidebar = ctk.CTkFrame(self, width=250)
        sidebar.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        sidebar.grid_propagate(False)
        
        ctk.CTkLabel(sidebar, text="Export Settings", font=ctk.CTkFont(size=16, weight="bold")).pack(pady=(15, 20), padx=15)
        
        ctk.CTkLabel(sidebar, text="Page Size:", anchor="w").pack(fill="x", padx=15, pady=(10, 5))
        size_menu = ctk.CTkOptionMenu(
            sidebar, 
            values=list(self.PAGE_SIZES.keys()),
            variable=self.page_size_var,
            command=lambda v: self._update_preview()
        )
        size_menu.pack(fill="x", padx=15)
        
        ctk.CTkLabel(sidebar, text="Fit Mode:", anchor="w").pack(fill="x", padx=15, pady=(15, 5))
        ctk.CTkRadioButton(sidebar, text="Fit to page", variable=self.fit_mode_var, value="fit", command=self._update_preview).pack(anchor="w", padx=15)
        ctk.CTkRadioButton(sidebar, text="Fill page (crop)", variable=self.fit_mode_var, value="fill", command=self._update_preview).pack(anchor="w", padx=15)
        ctk.CTkRadioButton(sidebar, text="Stretch to fit", variable=self.fit_mode_var, value="stretch", command=self._update_preview).pack(anchor="w", padx=15)
        ctk.CTkRadioButton(sidebar, text="Center (no scale)", variable=self.fit_mode_var, value="center", command=self._update_preview).pack(anchor="w", padx=15)
        
        ctk.CTkLabel(sidebar, text=f"Margin: {self.margin_var.get()}px", anchor="w").pack(fill="x", padx=15, pady=(15, 5))
        self.margin_label = sidebar.winfo_children()[-1]
        margin_slider = ctk.CTkSlider(sidebar, from_=0, to=100, variable=self.margin_var, command=self._on_margin_change)
        margin_slider.pack(fill="x", padx=15)
        
        ctk.CTkLabel(sidebar, text=f"Quality: {self.quality_var.get()}%", anchor="w").pack(fill="x", padx=15, pady=(15, 5))
        self.quality_label = sidebar.winfo_children()[-1]
        quality_slider = ctk.CTkSlider(sidebar, from_=50, to=100, variable=self.quality_var, command=self._on_quality_change)
        quality_slider.pack(fill="x", padx=15)
        
        ctk.CTkFrame(sidebar, height=2).pack(fill="x", padx=15, pady=20)
        
        info_text = f"{len(self.images)} pages"
        ctk.CTkLabel(sidebar, text=info_text, text_color="#888888").pack(padx=15)
        
        ctk.CTkFrame(sidebar, height=2).pack(fill="x", padx=15, pady=20)
        
        ctk.CTkButton(
            sidebar,
            text="Export PDF",
            fg_color="#28a745",
            hover_color="#218838",
            height=40,
            command=self._export
        ).pack(fill="x", padx=15, pady=5)
        
        ctk.CTkButton(
            sidebar,
            text="Cancel",
            fg_color="transparent",
            border_width=1,
            height=40,
            command=self.destroy
        ).pack(fill="x", padx=15, pady=5)
    
    def _on_margin_change(self, value):
        self.margin_label.configure(text=f"Margin: {int(value)}px")
        self._update_preview()
    
    def _on_quality_change(self, value):
        self.quality_label.configure(text=f"Quality: {int(value)}%")
    
    def _setup_preview(self):
        preview_frame = ctk.CTkFrame(self)
        preview_frame.grid(row=0, column=1, sticky="nsew", padx=(0, 10), pady=10)
        preview_frame.grid_columnconfigure(0, weight=1)
        preview_frame.grid_rowconfigure(1, weight=1)
        
        nav = ctk.CTkFrame(preview_frame, fg_color="transparent")
        nav.grid(row=0, column=0, sticky="ew", padx=10, pady=10)
        
        ctk.CTkButton(nav, text="< Prev", width=80, command=self._prev_page).pack(side="left")
        self.page_label = ctk.CTkLabel(nav, text="Page 1 / 1")
        self.page_label.pack(side="left", expand=True)
        ctk.CTkButton(nav, text="Next >", width=80, command=self._next_page).pack(side="right")
        
        self.canvas = tk.Canvas(preview_frame, bg="#404040", highlightthickness=1, highlightbackground="#666")
        self.canvas.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
    
    def _prev_page(self):
        if self.current_page > 0:
            self.current_page -= 1
            self._update_preview()
    
    def _next_page(self):
        if self.current_page < len(self.images) - 1:
            self.current_page += 1
            self._update_preview()
    
    def _update_preview(self):
        self.page_label.configure(text=f"Page {self.current_page + 1} / {len(self.images)}")
        
        self.canvas.update_idletasks()
        canvas_w = self.canvas.winfo_width() or 500
        canvas_h = self.canvas.winfo_height() or 600
        
        page_size = self.PAGE_SIZES.get(self.page_size_var.get())
        margin = self.margin_var.get()
        
        if page_size:
            page_w, page_h = page_size
        else:
            img = self.images[self.current_page]
            page_w, page_h = img.size
        
        scale = min((canvas_w - 40) / page_w, (canvas_h - 40) / page_h)
        preview_w = int(page_w * scale)
        preview_h = int(page_h * scale)
        
        preview_img = Image.new("RGB", (preview_w, preview_h), "#FFFFFF")
        
        img = self.images[self.current_page].copy()
        
        content_w = preview_w - int(margin * 2 * scale)
        content_h = preview_h - int(margin * 2 * scale)
        
        if content_w > 0 and content_h > 0:
            fit_mode = self.fit_mode_var.get()
            
            if fit_mode == "fit":
                img.thumbnail((content_w, content_h), Image.Resampling.LANCZOS)
            elif fit_mode == "fill":
                img_ratio = img.width / img.height
                content_ratio = content_w / content_h
                if img_ratio > content_ratio:
                    new_h = content_h
                    new_w = int(new_h * img_ratio)
                else:
                    new_w = content_w
                    new_h = int(new_w / img_ratio)
                img = img.resize((new_w, new_h), Image.Resampling.LANCZOS)
                left = (img.width - content_w) // 2
                top = (img.height - content_h) // 2
                img = img.crop((left, top, left + content_w, top + content_h))
            elif fit_mode == "stretch":
                img = img.resize((content_w, content_h), Image.Resampling.LANCZOS)
            
            x = (preview_w - img.width) // 2
            y = (preview_h - img.height) // 2
            preview_img.paste(img, (x, y))
        
        self.preview_tk = ImageTk.PhotoImage(preview_img)
        
        self.canvas.delete("all")
        x = (canvas_w - preview_w) // 2
        y = (canvas_h - preview_h) // 2
        self.canvas.create_rectangle(x-1, y-1, x+preview_w+1, y+preview_h+1, outline="#888", width=1)
        self.canvas.create_image(x, y, anchor="nw", image=self.preview_tk)
    
    def _export(self):
        filepath = filedialog.asksaveasfilename(
            title="Export PDF",
            defaultextension=".pdf",
            initialfile=f"{self.project_name}.pdf",
            filetypes=[("PDF files", "*.pdf")]
        )
        if not filepath:
            return
        
        try:
            page_size = self.PAGE_SIZES.get(self.page_size_var.get())
            margin = self.margin_var.get()
            quality = self.quality_var.get()
            fit_mode = self.fit_mode_var.get()
            
            pdf_images = []
            for img in self.images:
                if page_size:
                    page_w, page_h = page_size
                    page_img = Image.new("RGB", (page_w, page_h), "#FFFFFF")
                    
                    content_w = page_w - margin * 2
                    content_h = page_h - margin * 2
                    
                    fit_img = img.copy()
                    
                    if fit_mode == "fit":
                        fit_img.thumbnail((content_w, content_h), Image.Resampling.LANCZOS)
                    elif fit_mode == "fill":
                        img_ratio = fit_img.width / fit_img.height
                        content_ratio = content_w / content_h
                        if img_ratio > content_ratio:
                            new_h = content_h
                            new_w = int(new_h * img_ratio)
                        else:
                            new_w = content_w
                            new_h = int(new_w / img_ratio)
                        fit_img = fit_img.resize((new_w, new_h), Image.Resampling.LANCZOS)
                        left = (fit_img.width - content_w) // 2
                        top = (fit_img.height - content_h) // 2
                        fit_img = fit_img.crop((left, top, left + content_w, top + content_h))
                    elif fit_mode == "stretch":
                        fit_img = fit_img.resize((content_w, content_h), Image.Resampling.LANCZOS)
                    
                    x = (page_w - fit_img.width) // 2
                    y = (page_h - fit_img.height) // 2
                    page_img.paste(fit_img, (x, y))
                    pdf_images.append(page_img)
                else:
                    pdf_images.append(img.convert("RGB"))
            
            if pdf_images:
                pdf_images[0].save(
                    filepath, "PDF", 
                    save_all=True, 
                    append_images=pdf_images[1:],
                    quality=quality
                )
                messagebox.showinfo("Export Complete", f"Exported {len(pdf_images)} pages to:\n{filepath}")
                self.destroy()
        
        except Exception as e:
            messagebox.showerror("Error", f"Export failed: {e}")


class ProgressDialog(ctk.CTkToplevel):
    def __init__(self, parent, title: str, total: int):
        super().__init__(parent)
        
        self.title(title)
        self.geometry("400x120")
        self.resizable(False, False)
        self.transient(parent)
        
        self.total = total
        
        self.label = ctk.CTkLabel(self, text="Starting...")
        self.label.pack(pady=(20, 10))
        
        self.progress = ctk.CTkProgressBar(self, width=350)
        self.progress.pack(pady=10)
        self.progress.set(0)
        
        self.percent_label = ctk.CTkLabel(self, text="0%")
        self.percent_label.pack()
        
        self.update()
    
    def update_progress(self, current: int, message: str = ""):
        progress = current / self.total if self.total > 0 else 1
        self.progress.set(progress)
        self.percent_label.configure(text=f"{int(progress * 100)}%")
        if message:
            self.label.configure(text=message)
        self.update()


class ProjectKeybindHelpDialog(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        
        self.title("Project Keyboard Shortcuts")
        self.geometry("350x300")
        self.resizable(False, False)
        self.transient(parent)
        
        ctk.CTkLabel(
            self,
            text="Keyboard Shortcuts",
            font=ctk.CTkFont(size=18, weight="bold")
        ).pack(pady=(20, 15))
        
        shortcuts = [
            ("←  →", "Select previous/next page"),
            ("↑", "Move selected page up"),
            ("↓", "Move selected page down"),
            ("Delete / Backspace", "Remove selected page"),
            ("?", "Show this help"),
        ]
        
        text_frame = ctk.CTkFrame(self, fg_color="transparent")
        text_frame.pack(padx=20, pady=10, fill="both", expand=True)
        
        for key, desc in shortcuts:
            row = ctk.CTkFrame(text_frame, fg_color="transparent")
            row.pack(fill="x", pady=3)
            
            ctk.CTkLabel(
                row,
                text=key,
                font=ctk.CTkFont(family="Monaco", size=12),
                width=150,
                anchor="w"
            ).pack(side="left", padx=(10, 0))
            
            ctk.CTkLabel(
                row,
                text=desc,
                font=ctk.CTkFont(size=12),
                anchor="w"
            ).pack(side="left", padx=10)
        
        ctk.CTkButton(self, text="Close", command=self.destroy).pack(pady=15)
        
        self.bind("<Escape>", lambda e: self.destroy())
        self.focus()


class CommandDialog(ctk.CTkToplevel):
    def __init__(self, parent, execute_callback):
        super().__init__(parent)
        
        self.execute_callback = execute_callback
        
        self.title("")
        self.geometry("300x80")
        self.resizable(False, False)
        self.overrideredirect(True)
        
        x = parent.winfo_rootx() + parent.winfo_width() // 2 - 150
        y = parent.winfo_rooty() + parent.winfo_height() - 150
        self.geometry(f"300x80+{x}+{y}")
        
        self.transient(parent)
        self.grab_set()
        
        frame = ctk.CTkFrame(self, border_width=2, border_color="#1f538d")
        frame.pack(fill="both", expand=True, padx=2, pady=2)
        
        label = ctk.CTkLabel(frame, text=":", font=ctk.CTkFont(family="Monaco", size=18))
        label.pack(side="left", padx=(15, 5), pady=15)
        
        self.entry = ctk.CTkEntry(
            frame,
            font=ctk.CTkFont(family="Monaco", size=18),
            border_width=0,
            fg_color="transparent"
        )
        self.entry.pack(side="left", fill="x", expand=True, padx=(0, 15), pady=15)
        self.entry.focus()
        
        self.entry.bind("<Return>", self._submit)
        self.entry.bind("<Escape>", lambda e: self.destroy())
        self.bind("<FocusOut>", lambda e: self.destroy())
    
    def _submit(self, event=None):
        cmd = self.entry.get()
        self.destroy()
        self.execute_callback(cmd)


class KeybindHelpDialog(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        
        self.title("Keyboard Shortcuts")
        self.geometry("400x450")
        self.resizable(False, False)
        
        self.transient(parent)
        
        ctk.CTkLabel(
            self,
            text="Keyboard Shortcuts",
            font=ctk.CTkFont(size=18, weight="bold")
        ).pack(pady=(20, 15))
        
        shortcuts = [
            ("Navigation", [
                ("←  or  k", "Previous page"),
                ("→  or  j", "Next page"),
                ("g", "Go to first page"),
                ("G", "Go to last page"),
                (":14", "Go to page 14"),
                ("14 Enter", "Go to page 14"),
            ]),
            ("Actions", [
                ("c", "Copy current page"),
                ("?", "Show this help"),
            ]),
            ("Mouse", [
                ("Scroll", "Scroll page vertically"),
            ]),
        ]
        
        text_frame = ctk.CTkScrollableFrame(self, width=360, height=300)
        text_frame.pack(padx=20, pady=10, fill="both", expand=True)
        
        for section_title, bindings in shortcuts:
            ctk.CTkLabel(
                text_frame,
                text=section_title,
                font=ctk.CTkFont(size=14, weight="bold"),
                anchor="w"
            ).pack(fill="x", pady=(15, 5))
            
            for key, desc in bindings:
                row = ctk.CTkFrame(text_frame, fg_color="transparent")
                row.pack(fill="x", pady=2)
                
                ctk.CTkLabel(
                    row,
                    text=key,
                    font=ctk.CTkFont(family="Monaco", size=12),
                    width=120,
                    anchor="w"
                ).pack(side="left", padx=(10, 0))
                
                ctk.CTkLabel(
                    row,
                    text=desc,
                    font=ctk.CTkFont(size=12),
                    anchor="w"
                ).pack(side="left", padx=10)
        
        ctk.CTkButton(self, text="Close", command=self.destroy).pack(pady=15)
        
        self.bind("<Escape>", lambda e: self.destroy())
        self.bind("<question>", lambda e: self.destroy())
        self.focus()


class PDFCropToolApp(ctk.CTk):
    """Main application window."""
    
    def __init__(self):
        super().__init__()
        
        self.title("PDF Crop Tool")
        self.geometry("1200x800")
        self.minsize(900, 600)
        
        self.config = AppConfig()
        
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        
        self.current_frame: Optional[ctk.CTkFrame] = None
        self.show_welcome()
    
    def _show_frame(self, frame: ctk.CTkFrame):
        if self.current_frame:
            self.current_frame.destroy()
        
        self.current_frame = frame
        self.current_frame.grid(row=0, column=0, sticky="nsew")
    
    def show_welcome(self):
        self._show_frame(WelcomeScreen(self, self))
    
    def show_source_browser(self):
        self._show_frame(SourceBrowser(self, self))
    
    def show_source_editor(self, source: Source):
        self._show_frame(SourceEditor(self, self, source))
    
    def show_project_browser(self):
        self._show_frame(ProjectBrowser(self, self))
    
    def show_project_editor(self, project: Project):
        self._show_frame(ProjectEditor(self, self, project))


def main():
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("blue")
    
    app = PDFCropToolApp()
    app.mainloop()


if __name__ == "__main__":
    main()
