"""
ANC Character Count Report Generator - Desktop Application
-------------------------------------------------------
Professional, modern website-style desktop GUI for Swift Prosys.
Features:
  - Responsive full-width design without side gaps
  - Real-time batch-by-batch progress bar with percentage indicator
  - Full Dark and Light theme toggle with perfect contrast
  - Swift Prosys official branding (Swift-Prosys.png, Swift_Prosys.ico)
  - Quick-open buttons for generated Excel file & output folder
  - 100% core report generation logic in report_core.py preserved
"""

import os
import sys
import threading
import traceback
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
from PIL import Image

# Import business logic
from report_core import generate_report, generate_partial_link_report, generate_diff_reports
import updater

APP_VERSION = "1.0.0"  # bump this every release, and tag the GitHub Release the same way (v1.0.0)

# ---------------------------------------------------------------------------
# Path & Asset Resolution
# ---------------------------------------------------------------------------
BASE_DIR = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
LOGO_PATH = os.path.join(BASE_DIR, "Swift-Prosys.png")
ICON_PATH = os.path.join(BASE_DIR, "Swift_Prosys.ico")

# Appearance defaults
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


# ---------------------------------------------------------------------------
# Main Application Class
# ---------------------------------------------------------------------------
class DocTypeReportApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Window configuration
        self.title("ANC Character Count Report Generator — Swift Prosys")
        self.minsize(820, 580)
        self._center_window(1080, 720)

        # Window Icon
        if os.path.exists(ICON_PATH):
            try:
                self.iconbitmap(ICON_PATH)
            except Exception:
                pass

        # State Variables
        self.folder_path = ctk.StringVar(value="")
        self.project_id = ctk.StringVar(value="")
        self.progress_msg = ctk.StringVar(value="")
        self.progress_pct = ctk.StringVar(value="")
        self.last_generated_path = None
        self.current_theme = "dark"

        # Setup Layout
        self._setup_ui()

    def _center_window(self, width, height):
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        x = max(0, (sw - width) // 2)
        y = max(0, (sh - height) // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")

    # -----------------------------------------------------------------------
    # UI Construction
    # -----------------------------------------------------------------------
    def _setup_ui(self):
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # 1. Navigation Header Bar (Website Navbar style)
        self._build_header()

        # 2. Main Scrollable Container — smooth Canvas-based scroll (no lag)
        self._DARK_BG  = "#111827"   # matches header dark bg
        self._LIGHT_BG = "#FFFFFF"   # pure white for light mode

        self._scroll_canvas = tk.Canvas(
            self,
            highlightthickness=0,
            bg=self._DARK_BG         # starts in dark mode
        )
        self._scroll_canvas.grid(row=1, column=0, sticky="nsew", padx=0, pady=0)

        self._v_scrollbar = ctk.CTkScrollbar(
            self,
            command=self._scroll_canvas.yview
        )
        self._v_scrollbar.grid(row=1, column=1, sticky="ns")
        self._scroll_canvas.configure(yscrollcommand=self._v_scrollbar.set)

        self.main_scroll = ctk.CTkFrame(
            self._scroll_canvas,
            fg_color="transparent",
            corner_radius=0
        )
        self._canvas_window = self._scroll_canvas.create_window(
            (0, 0), window=self.main_scroll, anchor="nw"
        )

        # Resize canvas window when frame or canvas size changes
        self.main_scroll.bind("<Configure>", self._on_frame_configure)
        self._scroll_canvas.bind("<Configure>", self._on_canvas_configure)

        # Smooth scroll — only activate when mouse is over the scroll area
        # (avoids lag from bind_all that fires on every widget globally)
        self._scroll_canvas.bind("<Enter>", self._bind_mousewheel)
        self._scroll_canvas.bind("<Leave>", self._unbind_mousewheel)
        self.main_scroll.bind("<Enter>", self._bind_mousewheel)
        self.main_scroll.bind("<Leave>", self._unbind_mousewheel)

        self.main_scroll.grid_columnconfigure(0, weight=1)

        # Responsive full-width content wrapper (stretches with window)
        self.content_wrapper = ctk.CTkFrame(
            self.main_scroll,
            fg_color="transparent",
            corner_radius=0
        )
        self.content_wrapper.grid(row=0, column=0, sticky="nsew", padx=20, pady=16)
        self.content_wrapper.grid_columnconfigure(0, weight=1)

        # 3. Modern Segmented Tab Navigation
        self._build_tabs()

        # 4. Tab Views
        self.view_manual_key = ctk.CTkFrame(self.content_wrapper, fg_color="transparent")
        self.view_partially_linked = ctk.CTkFrame(self.content_wrapper, fg_color="transparent")

        self._build_manual_key_tab(self.view_manual_key)
        self._build_partially_linked_tab(self.view_partially_linked)

        # Show initial tab
        self.view_manual_key.grid(row=2, column=0, sticky="nsew", pady=(12, 20))

    # -----------------------------------------------------------------------
    # Header Navbar
    # -----------------------------------------------------------------------
    def _build_header(self):
        self.header_frame = ctk.CTkFrame(
            self,
            height=72,
            corner_radius=0,
            fg_color=("#FFFFFF", "#111827"),
            border_width=1,
            border_color=("#E5E7EB", "#1F2937")
        )
        self.header_frame.grid(row=0, column=0, sticky="ew")
        self.header_frame.grid_propagate(False)
        self.header_frame.grid_columnconfigure(1, weight=1)

        # Left: Brand Logo & Title
        brand_frame = ctk.CTkFrame(self.header_frame, fg_color="transparent")
        brand_frame.grid(row=0, column=0, padx=(20, 10), pady=12, sticky="w")

        # Company Logo
        if os.path.exists(LOGO_PATH):
            try:
                pil_logo = Image.open(LOGO_PATH)
                h = 36
                w = int(h * (pil_logo.width / pil_logo.height))
                self.logo_image = ctk.CTkImage(light_image=pil_logo, dark_image=pil_logo, size=(w, h))
                logo_lbl = ctk.CTkLabel(brand_frame, image=self.logo_image, text="")
                logo_lbl.pack(side="left", padx=(0, 12))
            except Exception:
                pass

        # Brand Text
        title_box = ctk.CTkFrame(brand_frame, fg_color="transparent")
        title_box.pack(side="left", fill="y")

        brand_title = ctk.CTkLabel(
            title_box,
            text="Indexing Report Toolkit",
            font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
            text_color=("#0F172A", "#F8FAFC")
        )
        brand_title.pack(anchor="w")

        brand_subtitle = ctk.CTkLabel(
            title_box,
            text="ANC Character Count Report Generator",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=("#64748B", "#94A3B8")
        )
        brand_subtitle.pack(anchor="w")

        # Right: Version Pill & Theme Switcher
        nav_right = ctk.CTkFrame(self.header_frame, fg_color="transparent")
        nav_right.grid(row=0, column=2, padx=20, pady=12, sticky="e")

        self.ver_badge = ctk.CTkLabel(
            nav_right,
            text=f" v{APP_VERSION}",
            font=ctk.CTkFont(family="Segoe UI", size=10, weight="bold"),
            text_color=("#0284C7", "#38BDF8"),
            fg_color=("#E0F2FE", "#0C4A6E"),
            corner_radius=12,
            padx=10,
            pady=4
        )
        self.ver_badge.pack(side="left", padx=(0, 8))

        self.update_btn = ctk.CTkButton(
            nav_right,
            text="🔄 Check for Updates",
            command=self.on_check_for_updates,
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            fg_color=("#F1F5F9", "#1E293B"),
            hover_color=("#E2E8F0", "#334155"),
            text_color=("#0F172A", "#F8FAFC"),
            corner_radius=16,
            height=30,
            width=140
        )
        self.update_btn.pack(side="left", padx=(0, 12))

        self.theme_segmented = ctk.CTkSegmentedButton(
            nav_right,
            values=["🌙 Dark", "☀️ Light"],
            command=self._on_theme_toggle,
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            selected_color="#2563EB",
            selected_hover_color="#1D4ED8",
            unselected_color=("#F1F5F9", "#1E293B"),
            unselected_hover_color=("#E2E8F0", "#334155"),
            text_color=("#0F172A", "#F8FAFC"),
            corner_radius=20,
            height=32
        )
        self.theme_segmented.set("🌙 Dark")
        self.theme_segmented.pack(side="left")

    # -----------------------------------------------------------------------
    # Segmented Navigation Tabs
    # -----------------------------------------------------------------------
    def _on_frame_configure(self, event=None):
        """Update scroll region when inner frame changes size."""
        self._scroll_canvas.configure(scrollregion=self._scroll_canvas.bbox("all"))

    def _on_canvas_configure(self, event):
        """Stretch inner frame to fill the full canvas width."""
        self._scroll_canvas.itemconfig(self._canvas_window, width=event.width)

    def _bind_mousewheel(self, event=None):
        """Activate scroll only when mouse enters the scroll area — prevents global lag."""
        self.bind_all("<MouseWheel>", self._on_mousewheel)

    def _unbind_mousewheel(self, event=None):
        """Deactivate scroll when mouse leaves scroll area."""
        self.unbind_all("<MouseWheel>")

    def _on_mousewheel(self, event):
        """Butter-smooth scroll: 4 units per wheel tick for fast, natural feel."""
        self._scroll_canvas.yview_scroll(int(-1 * (event.delta / 120)) * 4, "units")

    def _build_tabs(self):
        tab_container = ctk.CTkFrame(self.content_wrapper, fg_color="transparent")
        tab_container.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        tab_container.grid_columnconfigure(0, weight=1)

        self.tab_selector = ctk.CTkSegmentedButton(
            tab_container,
            values=["📁 Manual Key", "🔗 Keying Partially Linked"],
            command=self._on_tab_change,
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            selected_color="#2563EB",
            selected_hover_color="#1D4ED8",
            unselected_color=("#E2E8F0", "#1E293B"),
            unselected_hover_color=("#CBD5E1", "#334155"),
            text_color=("#0F172A", "#F8FAFC"),
            corner_radius=10,
            height=40
        )
        self.tab_selector.set("📁 Manual Key")
        self.tab_selector.pack(fill="x", padx=4)

    def _on_tab_change(self, selected_tab):
        if "Manual Key" in selected_tab:
            self.view_partially_linked.grid_forget()
            self.view_manual_key.grid(row=2, column=0, sticky="nsew", pady=(12, 20))
        else:
            self.view_manual_key.grid_forget()
            self.view_partially_linked.grid(row=2, column=0, sticky="nsew", pady=(12, 20))

    def _on_theme_toggle(self, value):
        if "Dark" in value:
            self.current_theme = "dark"
            ctk.set_appearance_mode("dark")
            self._apply_canvas_theme(self._DARK_BG)
        else:
            self.current_theme = "light"
            ctk.set_appearance_mode("light")
            self._apply_canvas_theme(self._LIGHT_BG)

    def _apply_canvas_theme(self, bg_color):
        """Force-sync the raw tk.Canvas and all transparent CTk children to the new bg."""
        self._scroll_canvas.configure(bg=bg_color)
        # Also explicitly set the inner frame so it doesn't bleed old color through
        self.main_scroll.configure(fg_color=bg_color)
        self.content_wrapper.configure(fg_color=bg_color)
        # Force Tk to repaint everything immediately
        self._scroll_canvas.update_idletasks()
        self.update_idletasks()

    # -----------------------------------------------------------------------
    # Tab 1: Manual Key (ANC Character Count Report Generator)
    # -----------------------------------------------------------------------
    def _build_manual_key_tab(self, parent):
        parent.grid_columnconfigure(0, weight=1)

        # 1. Hero Card: Description & Workflow Steps
        hero_card = ctk.CTkFrame(
            parent,
            corner_radius=14,
            fg_color=("#F8FAFC", "#1E293B"),
            border_width=1,
            border_color=("#E2E8F0", "#334155")
        )
        hero_card.grid(row=0, column=0, sticky="ew", pady=(0, 16))
        hero_card.grid_columnconfigure(0, weight=1)

        hero_inner = ctk.CTkFrame(hero_card, fg_color="transparent")
        hero_inner.pack(fill="both", expand=True, padx=22, pady=18)

        hero_title = ctk.CTkLabel(
            hero_inner,
            text="ANC Character Count Report Generator",
            font=ctk.CTkFont(family="Segoe UI", size=18, weight="bold"),
            text_color=("#0F172A", "#F8FAFC")
        )
        hero_title.pack(anchor="w")

        hero_desc = ctk.CTkLabel(
            hero_inner,
            text="Parses SQLite (.db3) batch databases, calculates page/record counts, and generates official "
                 "character counts.",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color=("#475569", "#94A3B8"),
            wraplength=1000,
            justify="left"
        )
        hero_desc.pack(anchor="w", pady=(4, 8))

        # Workflow Steps Pill Strip
        steps_strip = ctk.CTkFrame(hero_inner, fg_color="transparent")
        steps_strip.pack(anchor="w")

        steps = [
            ("1", "Select .db3 Folder"),
            ("2", "Confirm Project ID"),
            ("3", "Generate Excel Report")
        ]
        for idx, (num, text) in enumerate(steps):
            step_pill = ctk.CTkFrame(
                steps_strip,
                fg_color=("#E0F2FE", "#0F2942"),
                corner_radius=8,
                border_width=1,
                border_color=("#BAE6FD", "#1E4B70")
            )
            step_pill.pack(side="left", padx=(0, 10))

            ctk.CTkLabel(
                step_pill,
                text=f" {num} ",
                font=ctk.CTkFont(family="Segoe UI", size=10, weight="bold"),
                text_color="white",
                fg_color="#0284C7",
                corner_radius=6,
                padx=5, pady=1
            ).pack(side="left", padx=5, pady=4)

            ctk.CTkLabel(
                step_pill,
                text=text,
                font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
                text_color=("#0369A1", "#7DD3FC")
            ).pack(side="left", padx=(0, 8), pady=4)

        # 2. Configuration Card (Inputs)
        config_card = ctk.CTkFrame(
            parent,
            corner_radius=14,
            fg_color=("#FFFFFF", "#1E293B"),
            border_width=1,
            border_color=("#E2E8F0", "#334155")
        )
        config_card.grid(row=1, column=0, sticky="ew", pady=(0, 16))
        config_card.grid_columnconfigure(0, weight=1)

        cfg_inner = ctk.CTkFrame(config_card, fg_color="transparent")
        cfg_inner.pack(fill="both", expand=True, padx=24, pady=22)
        cfg_inner.grid_columnconfigure(0, weight=1)

        # --- DB3 Folder Selection ---
        folder_header = ctk.CTkFrame(cfg_inner, fg_color="transparent")
        folder_header.grid(row=0, column=0, sticky="w", pady=(0, 6))

        ctk.CTkLabel(
            folder_header,
            text="Database Source Folder",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            text_color=("#0F172A", "#F8FAFC")
        ).pack(side="left")

        ctk.CTkLabel(
            folder_header,
            text=" (Contains batch .db3 files)",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=("#64748B", "#94A3B8")
        ).pack(side="left")

        folder_row = ctk.CTkFrame(cfg_inner, fg_color="transparent")
        folder_row.grid(row=1, column=0, sticky="ew")
        folder_row.grid_columnconfigure(0, weight=1)

        self.entry_folder = ctk.CTkEntry(
            folder_row,
            textvariable=self.folder_path,
            placeholder_text="Click 'Browse Folder' or paste the absolute path to your .db3 folder...",
            height=44,
            corner_radius=8,
            font=ctk.CTkFont(family="Segoe UI", size=11),
            fg_color=("#F8FAFC", "#0F172A"),
            border_color=("#CBD5E1", "#475569"),
            text_color=("#0F172A", "#F8FAFC")
        )
        self.entry_folder.grid(row=0, column=0, sticky="ew", padx=(0, 12))

        self.btn_browse = ctk.CTkButton(
            folder_row,
            text="📂 Browse Folder",
            command=self.on_browse_folder,
            height=44,
            corner_radius=8,
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            fg_color="#0284C7",
            hover_color="#0369A1"
        )
        self.btn_browse.grid(row=0, column=1)

        # Real-time Folder Batch Stats Chip
        self.chip_frame = ctk.CTkFrame(cfg_inner, fg_color="transparent")
        self.chip_frame.grid(row=2, column=0, sticky="w", pady=(8, 14))

        self.chip_files = ctk.CTkLabel(
            self.chip_frame,
            text="📦 No folder selected",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            text_color=("#475569", "#94A3B8"),
            fg_color=("#F1F5F9", "#0F172A"),
            corner_radius=6,
            padx=10,
            pady=3
        )
        self.chip_files.pack(side="left")

        # Subtle Horizontal Separator
        ctk.CTkFrame(
            cfg_inner,
            height=1,
            fg_color=("#E2E8F0", "#334155")
        ).grid(row=3, column=0, sticky="ew", pady=(4, 16))

        # --- Project ID Field ---
        ctk.CTkLabel(
            cfg_inner,
            text="Project Identification (Project ID)",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            text_color=("#0F172A", "#F8FAFC")
        ).grid(row=4, column=0, sticky="w", pady=(0, 6))

        pid_row = ctk.CTkFrame(cfg_inner, fg_color="transparent")
        pid_row.grid(row=5, column=0, sticky="w")

        self.entry_pid = ctk.CTkEntry(
            pid_row,
            textvariable=self.project_id,
            placeholder_text="e.g. 62884.IDX.001",
            width=360,
            height=44,
            corner_radius=8,
            font=ctk.CTkFont(family="Segoe UI", size=12),
            fg_color=("#F8FAFC", "#0F172A"),
            border_color=("#CBD5E1", "#475569"),
            text_color=("#0F172A", "#F8FAFC")
        )
        self.entry_pid.pack(side="left", padx=(0, 10))

        ctk.CTkLabel(
            pid_row,
            text="✨ Auto-extracted from folder name (Editable)",
            font=ctk.CTkFont(family="Segoe UI", size=11, slant="italic"),
            text_color=("#64748B", "#94A3B8")
        ).pack(side="left")

        # 3. Action Execution & Real-Time Progress Card
        action_card = ctk.CTkFrame(
            parent,
            corner_radius=14,
            fg_color=("#FFFFFF", "#1E293B"),
            border_width=1,
            border_color=("#E2E8F0", "#334155")
        )
        action_card.grid(row=2, column=0, sticky="ew", pady=(0, 16))
        action_card.grid_columnconfigure(0, weight=1)

        act_inner = ctk.CTkFrame(action_card, fg_color="transparent")
        act_inner.pack(fill="both", expand=True, padx=24, pady=20)
        act_inner.grid_columnconfigure(0, weight=1)

        # Primary Call to Action Button
        self.btn_generate = ctk.CTkButton(
            act_inner,
            text="⚡ Generate Report (.xlsx)",
            command=self.on_generate,
            height=50,
            corner_radius=10,
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            fg_color="#2563EB",
            hover_color="#1D4ED8"
        )
        self.btn_generate.grid(row=0, column=0, sticky="ew")

        # Real-time Status Text & Percentage Bar Container
        self.progress_container = ctk.CTkFrame(act_inner, fg_color="transparent")
        self.progress_container.grid(row=1, column=0, sticky="ew", pady=(16, 0))
        self.progress_container.grid_columnconfigure(0, weight=1)

        # Header of progress: Status text on left, Percentage on right
        prog_header = ctk.CTkFrame(self.progress_container, fg_color="transparent")
        prog_header.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        prog_header.grid_columnconfigure(0, weight=1)

        self.lbl_progress_status = ctk.CTkLabel(
            prog_header,
            textvariable=self.progress_msg,
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color=("#0F172A", "#F8FAFC"),
            anchor="w"
        )
        self.lbl_progress_status.grid(row=0, column=0, sticky="w")

        self.lbl_progress_pct = ctk.CTkLabel(
            prog_header,
            textvariable=self.progress_pct,
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color=("#0284C7", "#38BDF8"),
            anchor="e"
        )
        self.lbl_progress_pct.grid(row=0, column=1, sticky="e")

        # Real-time Progress Bar (0.0 to 1.0)
        self.progress_bar = ctk.CTkProgressBar(
            self.progress_container,
            height=10,
            corner_radius=5,
            fg_color=("#E2E8F0", "#334155"),
            progress_color="#2563EB"
        )
        self.progress_bar.grid(row=1, column=0, sticky="ew")
        self.progress_bar.set(0)

        # Hide progress container initially until user clicks Generate or selects folder
        self.progress_container.grid_remove()

        # Success Action Buttons (Shown when export completes)
        self.success_action_row = ctk.CTkFrame(act_inner, fg_color="transparent")
        self.success_action_row.grid(row=2, column=0, sticky="w", pady=(14, 0))

        self.btn_open_file = ctk.CTkButton(
            self.success_action_row,
            text="📊 Open Generated Excel File",
            command=self._open_excel_file,
            height=38,
            corner_radius=8,
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            fg_color="#16A34A",
            hover_color="#15803D"
        )

        self.btn_open_folder = ctk.CTkButton(
            self.success_action_row,
            text="📂 Open Output Folder",
            command=self._open_output_folder,
            height=38,
            corner_radius=8,
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            fg_color=("#475569", "#334155"),
            hover_color=("#334155", "#475569")
        )

    # -----------------------------------------------------------------------
    # Tab 2: Keying Partially Linked (Old + New two-pass report)
    # -----------------------------------------------------------------------
    def _build_partially_linked_tab(self, parent):
        parent.grid_columnconfigure(0, weight=1)

        # 1. Hero Card
        hero_card = ctk.CTkFrame(
            parent,
            corner_radius=14,
            fg_color=("#F8FAFC", "#1E293B"),
            border_width=1,
            border_color=("#E2E8F0", "#334155")
        )
        hero_card.grid(row=0, column=0, sticky="ew", pady=(0, 16))
        hero_card.grid_columnconfigure(0, weight=1)

        hero_inner = ctk.CTkFrame(hero_card, fg_color="transparent")
        hero_inner.pack(fill="both", expand=True, padx=22, pady=18)

        ctk.CTkLabel(
            hero_inner,
            text="Keying Partially Linked — Two-Pass Report",
            font=ctk.CTkFont(family="Segoe UI", size=18, weight="bold"),
            text_color=("#0F172A", "#F8FAFC")
        ).pack(anchor="w")

        ctk.CTkLabel(
            hero_inner,
            text="Customer Data = Partially Linked (name-linking only) batches. SPS Data = same batches after full keying.\n"
                 "Character Count automatically excludes whichever field(s) were already counted\n"
                 "during Partial Linking (auto-detected from the Customer Data folder), so nothing is double-counted.",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color=("#475569", "#94A3B8"),
            wraplength=1000,
            justify="left"
        ).pack(anchor="w", pady=(4, 8))

        # 2. Configuration Card (Old + New folder pickers, Project ID)
        config_card = ctk.CTkFrame(
            parent,
            corner_radius=14,
            fg_color=("#FFFFFF", "#1E293B"),
            border_width=1,
            border_color=("#E2E8F0", "#334155")
        )
        config_card.grid(row=1, column=0, sticky="ew", pady=(0, 16))
        config_card.grid_columnconfigure(0, weight=1)

        cfg_inner = ctk.CTkFrame(config_card, fg_color="transparent")
        cfg_inner.pack(fill="both", expand=True, padx=24, pady=22)
        cfg_inner.grid_columnconfigure(0, weight=1)

        self.pl_old_folder = ctk.StringVar(value="")
        self.pl_new_folder = ctk.StringVar(value="")
        self.pl_project_id = ctk.StringVar(value="")

        # --- Old folder ---
        ctk.CTkLabel(
            cfg_inner, text="Customer Data (Partial Link .db3 files)",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            text_color=("#0F172A", "#F8FAFC")
        ).grid(row=0, column=0, sticky="w", pady=(0, 6))

        old_row = ctk.CTkFrame(cfg_inner, fg_color="transparent")
        old_row.grid(row=1, column=0, sticky="ew", pady=(0, 6))
        old_row.grid_columnconfigure(0, weight=1)

        self.pl_entry_old = ctk.CTkEntry(
            old_row, textvariable=self.pl_old_folder,
            placeholder_text="Click 'Browse' or paste the absolute path to the Customer Data .db3 folder...",
            height=44, corner_radius=8, font=ctk.CTkFont(family="Segoe UI", size=11),
            fg_color=("#F8FAFC", "#0F172A"), border_color=("#CBD5E1", "#475569"),
            text_color=("#0F172A", "#F8FAFC")
        )
        self.pl_entry_old.grid(row=0, column=0, sticky="ew", padx=(0, 12))

        ctk.CTkButton(
            old_row, text="📂 Browse Customer Data", command=self.on_browse_old_folder,
            height=44, corner_radius=8, font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            fg_color="#0284C7", hover_color="#0369A1"
        ).grid(row=0, column=1)

        self.pl_chip_old = ctk.CTkLabel(
            cfg_inner, text="📦 No Customer folder selected",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            text_color=("#475569", "#94A3B8"), fg_color=("#F1F5F9", "#0F172A"),
            corner_radius=6, padx=10, pady=3
        )
        self.pl_chip_old.grid(row=2, column=0, sticky="w", pady=(0, 14))

        # --- New folder ---
        ctk.CTkLabel(
            cfg_inner, text="SPS Data (fully-keyed .db3 files)",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            text_color=("#0F172A", "#F8FAFC")
        ).grid(row=3, column=0, sticky="w", pady=(0, 6))

        new_row = ctk.CTkFrame(cfg_inner, fg_color="transparent")
        new_row.grid(row=4, column=0, sticky="ew", pady=(0, 6))
        new_row.grid_columnconfigure(0, weight=1)

        self.pl_entry_new = ctk.CTkEntry(
            new_row, textvariable=self.pl_new_folder,
            placeholder_text="Click 'Browse' or paste the absolute path to the SPS Data .db3 folder...",
            height=44, corner_radius=8, font=ctk.CTkFont(family="Segoe UI", size=11),
            fg_color=("#F8FAFC", "#0F172A"), border_color=("#CBD5E1", "#475569"),
            text_color=("#0F172A", "#F8FAFC")
        )
        self.pl_entry_new.grid(row=0, column=0, sticky="ew", padx=(0, 12))

        ctk.CTkButton(
            new_row, text="📂 Browse SPS Data", command=self.on_browse_new_folder,
            height=44, corner_radius=8, font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            fg_color="#0284C7", hover_color="#0369A1"
        ).grid(row=0, column=1)

        self.pl_chip_new = ctk.CTkLabel(
            cfg_inner, text="📦 No SPS Data folder selected",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            text_color=("#475569", "#94A3B8"), fg_color=("#F1F5F9", "#0F172A"),
            corner_radius=6, padx=10, pady=3
        )
        self.pl_chip_new.grid(row=5, column=0, sticky="w", pady=(0, 14))

        ctk.CTkFrame(
            cfg_inner, height=1, fg_color=("#E2E8F0", "#334155")
        ).grid(row=6, column=0, sticky="ew", pady=(4, 16))

        # --- Project ID ---
        ctk.CTkLabel(
            cfg_inner, text="Project Identification (Project ID)",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            text_color=("#0F172A", "#F8FAFC")
        ).grid(row=7, column=0, sticky="w", pady=(0, 6))

        pid_row = ctk.CTkFrame(cfg_inner, fg_color="transparent")
        pid_row.grid(row=8, column=0, sticky="w")

        self.pl_entry_pid = ctk.CTkEntry(
            pid_row, textvariable=self.pl_project_id,
            placeholder_text="e.g. 62105.IDX.008", width=360, height=44, corner_radius=8,
            font=ctk.CTkFont(family="Segoe UI", size=12),
            fg_color=("#F8FAFC", "#0F172A"), border_color=("#CBD5E1", "#475569"),
            text_color=("#0F172A", "#F8FAFC")
        )
        self.pl_entry_pid.pack(side="left", padx=(0, 10))

        ctk.CTkLabel(
            pid_row, text="✨ Auto-extracted from SPS Data folder name (Editable)",
            font=ctk.CTkFont(family="Segoe UI", size=11, slant="italic"),
            text_color=("#64748B", "#94A3B8")
        ).pack(side="left")

        # 3. Action + progress card
        action_card = ctk.CTkFrame(
            parent, corner_radius=14, fg_color=("#FFFFFF", "#1E293B"),
            border_width=1, border_color=("#E2E8F0", "#334155")
        )
        action_card.grid(row=2, column=0, sticky="ew", pady=(0, 16))
        action_card.grid_columnconfigure(0, weight=1)

        act_inner = ctk.CTkFrame(action_card, fg_color="transparent")
        act_inner.pack(fill="both", expand=True, padx=24, pady=20)
        act_inner.grid_columnconfigure(0, weight=1)

        self.pl_btn_generate = ctk.CTkButton(
            act_inner, text="⚡ Generate Partially-Linked Report (.xlsx)",
            command=self.on_generate_partial_link, height=50, corner_radius=10,
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            fg_color="#2563EB", hover_color="#1D4ED8"
        )
        self.pl_btn_generate.grid(row=0, column=0, sticky="ew")

        self.pl_btn_generate_diff = ctk.CTkButton(
            act_inner, text="🔍 Generate Comparison / Audit Reports (one file per batch)",
            command=self.on_generate_diff_reports, height=44, corner_radius=10,
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            fg_color=("#7C3AED", "#6D28D9"), hover_color=("#6D28D9", "#5B21B6")
        )
        self.pl_btn_generate_diff.grid(row=1, column=0, sticky="ew", pady=(10, 0))

        ctk.CTkLabel(
            act_inner,
            text="Shows Customer Data vs SPS Data side-by-side per field, with new/additional characters highlighted in red —\n"
                 "like an XLCompare diff — so you can see exactly what's driving each batch's Character Count.",
            font=ctk.CTkFont(family="Segoe UI", size=11), text_color=("#64748B", "#94A3B8"),
            justify="left"
        ).grid(row=2, column=0, sticky="w", pady=(6, 0))

        self.pl_progress_msg = ctk.StringVar(value="")
        self.pl_progress_pct = ctk.StringVar(value="")

        self.pl_progress_container = ctk.CTkFrame(act_inner, fg_color="transparent")
        self.pl_progress_container.grid(row=3, column=0, sticky="ew", pady=(16, 0))
        self.pl_progress_container.grid_columnconfigure(0, weight=1)

        prog_header = ctk.CTkFrame(self.pl_progress_container, fg_color="transparent")
        prog_header.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        prog_header.grid_columnconfigure(0, weight=1)

        self.pl_lbl_progress_status = ctk.CTkLabel(
            prog_header, textvariable=self.pl_progress_msg,
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color=("#0F172A", "#F8FAFC"), anchor="w"
        )
        self.pl_lbl_progress_status.grid(row=0, column=0, sticky="w")

        ctk.CTkLabel(
            prog_header, textvariable=self.pl_progress_pct,
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color=("#0284C7", "#38BDF8"), anchor="e"
        ).grid(row=0, column=1, sticky="e")

        self.pl_progress_bar = ctk.CTkProgressBar(
            self.pl_progress_container, height=10, corner_radius=5,
            fg_color=("#E2E8F0", "#334155"), progress_color="#2563EB"
        )
        self.pl_progress_bar.grid(row=1, column=0, sticky="ew")
        self.pl_progress_bar.set(0)
        self.pl_progress_container.grid_remove()

        self.pl_success_action_row = ctk.CTkFrame(act_inner, fg_color="transparent")
        self.pl_success_action_row.grid(row=4, column=0, sticky="w", pady=(14, 0))

        self.pl_btn_open_file = ctk.CTkButton(
            self.pl_success_action_row, text="📊 Open Generated Excel File",
            command=self._pl_open_excel_file, height=38, corner_radius=8,
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            fg_color="#16A34A", hover_color="#15803D"
        )
        self.pl_btn_open_folder = ctk.CTkButton(
            self.pl_success_action_row, text="📂 Open Output Folder",
            command=self._pl_open_output_folder, height=38, corner_radius=8,
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            fg_color=("#475569", "#334155"), hover_color=("#334155", "#475569")
        )
        self.pl_last_generated_path = None

    # -----------------------------------------------------------------------
    # Business Logic Event Handlers
    # -----------------------------------------------------------------------
    def on_browse_folder(self):
        folder = filedialog.askdirectory(title="Select Folder with .db3 Database Batches")
        if folder:
            self.folder_path.set(folder)
            detected_pid = os.path.basename(folder.rstrip("/\\"))
            self.project_id.set(detected_pid)

            # Count db3 files
            db3_files = [f for f in os.listdir(folder) if f.lower().endswith(".db3")]
            file_count = len(db3_files)

            if file_count > 0:
                self.chip_files.configure(
                    text=f"📦 {file_count} .db3 Database Batch(es) Found",
                    text_color="#16A34A",
                    fg_color=("#DCFCE7", "#052E16")
                )
            else:
                self.chip_files.configure(
                    text="⚠️ 0 .db3 files found in folder",
                    text_color="#DC2626",
                    fg_color=("#FEE2E2", "#450A0A")
                )

    def on_generate(self):
        folder = self.folder_path.get().strip()
        if not folder or not os.path.isdir(folder):
            messagebox.showerror("Invalid Folder", "Please browse and select a valid folder containing .db3 files.")
            return

        db3_files = [f for f in os.listdir(folder) if f.lower().endswith(".db3")]
        if not db3_files:
            messagebox.showerror("No Batches Found", "No .db3 database files found in the selected folder.")
            return

        project_id = self.project_id.get().strip() or os.path.basename(folder.rstrip("/\\"))
        default_filename = f"{project_id}_Project Report_DataExport.xlsx"

        out_path = filedialog.asksaveasfilename(
            title="Save Report As",
            defaultextension=".xlsx",
            initialfile=default_filename,
            filetypes=[("Excel Workbook", "*.xlsx")]
        )
        if not out_path:
            return

        # Prepare UI for real-time progress
        self.btn_generate.configure(state="disabled")
        self.btn_browse.configure(state="disabled")

        # Hide any previous quick buttons
        self.btn_open_file.pack_forget()
        self.btn_open_folder.pack_forget()

        # Show progress bar & labels
        self.progress_container.grid()
        self.progress_bar.set(0)
        self.progress_msg.set(f"Initializing extraction of {len(db3_files)} batch(es)...")
        self.progress_pct.set("0%")
        self.lbl_progress_status.configure(text_color=("#0F172A", "#F8FAFC"))

        # Launch worker thread with real-time callback
        threading.Thread(
            target=self._worker_thread,
            args=(folder, out_path, project_id),
            daemon=True
        ).start()

    def _worker_thread(self, folder, out_path, project_id):
        def progress_cb(current, total, msg):
            self.after(0, self._update_progress, current, total, msg)

        try:
            path, n_rows = generate_report(
                folder, out_path, project_id=project_id, progress_callback=progress_cb
            )
            self.after(0, self._on_generate_success, path, n_rows)
        except Exception as e:
            full_trace = traceback.format_exc()
            self.after(0, self._on_generate_error, str(e), full_trace)

    def _update_progress(self, current, total, msg):
        fraction = (current / total) if total > 0 else 0
        pct = int(fraction * 100)
        self.progress_bar.set(fraction)
        self.progress_pct.set(f"{pct}%")
        self.progress_msg.set(msg)

    def _on_generate_success(self, path, n_rows):
        self.progress_bar.set(1.0)
        self.progress_pct.set("100%")
        self.progress_msg.set(f"✅ Success: Wrote {n_rows} rows to {os.path.basename(path)}")
        self.lbl_progress_status.configure(text_color="#16A34A")

        self.btn_generate.configure(state="normal")
        self.btn_browse.configure(state="normal")
        self.last_generated_path = path

        # Show quick access buttons
        self.btn_open_file.pack(side="left", padx=(0, 10))
        self.btn_open_folder.pack(side="left")

        messagebox.showinfo(
            "Report Complete",
            f"Report generated successfully!\n\nLocation:\n{path}"
        )

    def _on_generate_error(self, err_msg, full_trace):
        self.progress_bar.set(0)
        self.progress_pct.set("")
        self.progress_msg.set(f"❌ Error: {err_msg}")
        self.lbl_progress_status.configure(text_color="#DC2626")

        self.btn_generate.configure(state="normal")
        self.btn_browse.configure(state="normal")

        messagebox.showerror("Generation Error", f"Failed to generate report:\n\n{err_msg}")
        print(full_trace, file=sys.stderr)

    def _open_excel_file(self):
        if self.last_generated_path and os.path.exists(self.last_generated_path):
            try:
                os.startfile(self.last_generated_path)
            except Exception as e:
                messagebox.showerror("Cannot Open File", str(e))

    def _open_output_folder(self):
        if self.last_generated_path:
            folder = os.path.dirname(self.last_generated_path)
            if os.path.exists(folder):
                try:
                    os.startfile(folder)
                except Exception as e:
                    messagebox.showerror("Cannot Open Folder", str(e))

    # -----------------------------------------------------------------------
    # Tab 2 Event Handlers: Keying Partially Linked
    # -----------------------------------------------------------------------
    def on_browse_old_folder(self):
        folder = filedialog.askdirectory(title="Select Customer Data (Partially Linked) .db3 Folder")
        if folder:
            self.pl_old_folder.set(folder)
            db3_files = [f for f in os.listdir(folder) if f.lower().endswith(".db3")]
            if db3_files:
                self.pl_chip_old.configure(
                    text=f"📦 {len(db3_files)} .db3 file(s) found",
                    text_color="#16A34A", fg_color=("#DCFCE7", "#052E16")
                )
            else:
                self.pl_chip_old.configure(
                    text="⚠️ 0 .db3 files found in folder",
                    text_color="#DC2626", fg_color=("#FEE2E2", "#450A0A")
                )

    def on_browse_new_folder(self):
        folder = filedialog.askdirectory(title="Select SPS Data (fully-keyed) .db3 Folder")
        if folder:
            self.pl_new_folder.set(folder)
            self.pl_project_id.set(os.path.basename(folder.rstrip("/\\")))
            db3_files = [f for f in os.listdir(folder) if f.lower().endswith(".db3")]
            if db3_files:
                self.pl_chip_new.configure(
                    text=f"📦 {len(db3_files)} .db3 file(s) found",
                    text_color="#16A34A", fg_color=("#DCFCE7", "#052E16")
                )
            else:
                self.pl_chip_new.configure(
                    text="⚠️ 0 .db3 files found in folder",
                    text_color="#DC2626", fg_color=("#FEE2E2", "#450A0A")
                )

    def on_generate_partial_link(self):
        old_folder = self.pl_old_folder.get().strip()
        new_folder = self.pl_new_folder.get().strip()

        if not old_folder or not os.path.isdir(old_folder):
            messagebox.showerror("Invalid Folder", "Please browse and select a valid Customer Data folder.")
            return
        if not new_folder or not os.path.isdir(new_folder):
            messagebox.showerror("Invalid Folder", "Please browse and select a valid SPS Data folder.")
            return

        new_db3 = [f for f in os.listdir(new_folder) if f.lower().endswith(".db3")]
        if not new_db3:
            messagebox.showerror("No Batches Found", "No .db3 database files found in the SPS Data folder.")
            return

        project_id = self.pl_project_id.get().strip() or os.path.basename(new_folder.rstrip("/\\"))
        default_filename = f"{project_id}_Project Report_PartiallyLinked.xlsx"

        out_path = filedialog.asksaveasfilename(
            title="Save Partially-Linked Report As",
            defaultextension=".xlsx",
            initialfile=default_filename,
            filetypes=[("Excel Workbook", "*.xlsx")]
        )
        if not out_path:
            return

        self.pl_btn_generate.configure(state="disabled")
        self.pl_btn_generate_diff.configure(state="disabled")
        self.pl_btn_open_file.pack_forget()
        self.pl_btn_open_folder.pack_forget()

        self.pl_progress_container.grid()
        self.pl_progress_bar.set(0)
        self.pl_progress_msg.set(f"Initializing extraction of {len(new_db3)} batch(es)...")
        self.pl_progress_pct.set("0%")
        self.pl_lbl_progress_status.configure(text_color=("#0F172A", "#F8FAFC"))

        threading.Thread(
            target=self._pl_worker_thread,
            args=(old_folder, new_folder, out_path, project_id),
            daemon=True
        ).start()

    def _pl_worker_thread(self, old_folder, new_folder, out_path, project_id):
        def progress_cb(current, total, msg):
            self.after(0, self._pl_update_progress, current, total, msg)

        try:
            path, n_rows = generate_partial_link_report(
                old_folder, new_folder, out_path,
                project_id=project_id, progress_callback=progress_cb
            )
            self.after(0, self._pl_on_generate_success, path, n_rows)
        except Exception as e:
            full_trace = traceback.format_exc()
            self.after(0, self._pl_on_generate_error, str(e), full_trace)

    def _pl_update_progress(self, current, total, msg):
        fraction = (current / total) if total > 0 else 0
        self.pl_progress_bar.set(fraction)
        self.pl_progress_pct.set(f"{int(fraction * 100)}%")
        self.pl_progress_msg.set(msg)

    def _pl_on_generate_success(self, path, n_rows):
        self.pl_progress_bar.set(1.0)
        self.pl_progress_pct.set("100%")
        self.pl_progress_msg.set(f"✅ Success: Wrote {n_rows} rows to {os.path.basename(path)}")
        self.pl_lbl_progress_status.configure(text_color="#16A34A")

        self.pl_btn_generate.configure(state="normal")
        self.pl_btn_generate_diff.configure(state="normal")
        self.pl_last_generated_path = path

        self.pl_btn_open_file.pack(side="left", padx=(0, 10))
        self.pl_btn_open_folder.pack(side="left")

        messagebox.showinfo(
            "Report Complete",
            f"Partially-Linked Report generated successfully!\n\nLocation:\n{path}"
        )

    def _pl_on_generate_error(self, err_msg, full_trace):
        self.pl_progress_bar.set(0)
        self.pl_progress_pct.set("")
        self.pl_progress_msg.set(f"❌ Error: {err_msg}")
        self.pl_lbl_progress_status.configure(text_color="#DC2626")

        self.pl_btn_generate.configure(state="normal")
        self.pl_btn_generate_diff.configure(state="normal")

        messagebox.showerror("Generation Error", f"Failed to generate report:\n\n{err_msg}")
        print(full_trace, file=sys.stderr)

    def _pl_open_excel_file(self):
        if self.pl_last_generated_path and os.path.exists(self.pl_last_generated_path):
            try:
                os.startfile(self.pl_last_generated_path)
            except Exception as e:
                messagebox.showerror("Cannot Open File", str(e))

    def _pl_open_output_folder(self):
        if self.pl_last_generated_path:
            folder = os.path.dirname(self.pl_last_generated_path) if os.path.isfile(self.pl_last_generated_path) else self.pl_last_generated_path
            if os.path.exists(folder):
                try:
                    os.startfile(folder)
                except Exception as e:
                    messagebox.showerror("Cannot Open Folder", str(e))

    # -----------------------------------------------------------------------
    # Tab 2: Comparison / Audit Reports (one file per batch)
    # -----------------------------------------------------------------------
    def on_generate_diff_reports(self):
        old_folder = self.pl_old_folder.get().strip()
        new_folder = self.pl_new_folder.get().strip()

        if not old_folder or not os.path.isdir(old_folder):
            messagebox.showerror("Invalid Folder", "Please browse and select a valid Customer Data folder.")
            return
        if not new_folder or not os.path.isdir(new_folder):
            messagebox.showerror("Invalid Folder", "Please browse and select a valid SPS Data folder.")
            return

        output_folder = filedialog.askdirectory(title="Select OUTPUT folder for the audit reports")
        if not output_folder:
            return

        self.pl_btn_generate.configure(state="disabled")
        self.pl_btn_generate_diff.configure(state="disabled")
        self.pl_btn_open_file.pack_forget()
        self.pl_btn_open_folder.pack_forget()

        self.pl_progress_container.grid()
        self.pl_progress_bar.set(0)
        self.pl_progress_msg.set("Initializing comparison...")
        self.pl_progress_pct.set("0%")
        self.pl_lbl_progress_status.configure(text_color=("#0F172A", "#F8FAFC"))

        threading.Thread(
            target=self._diff_worker_thread,
            args=(old_folder, new_folder, output_folder),
            daemon=True
        ).start()

    def _diff_worker_thread(self, old_folder, new_folder, output_folder):
        def progress_cb(current, total, msg):
            self.after(0, self._pl_update_progress, current, total, msg)

        try:
            out_folder, files = generate_diff_reports(
                old_folder, new_folder, output_folder, progress_callback=progress_cb
            )
            self.after(0, self._diff_on_success, out_folder, files)
        except Exception as e:
            full_trace = traceback.format_exc()
            self.after(0, self._pl_on_generate_error, str(e), full_trace)

    def _diff_on_success(self, out_folder, files):
        self.pl_progress_bar.set(1.0)
        self.pl_progress_pct.set("100%")
        self.pl_progress_msg.set(f"✅ Success: Wrote {len(files)} comparison file(s) to {out_folder}")
        self.pl_lbl_progress_status.configure(text_color="#16A34A")

        self.pl_btn_generate.configure(state="normal")
        self.pl_btn_generate_diff.configure(state="normal")
        self.pl_last_generated_path = out_folder

        self.pl_btn_open_folder.configure(text="📂 Open Comparison Reports Folder")
        self.pl_btn_open_folder.pack(side="left")

        messagebox.showinfo(
            "Comparison Reports Complete",
            f"{len(files)} audit report(s) generated!\n\nFolder:\n{out_folder}"
        )

    # -----------------------------------------------------------------------
    # Self-Update ("Check for Updates" button, GitHub Releases-backed)
    # -----------------------------------------------------------------------
    def on_check_for_updates(self):
        self.update_btn.configure(state="disabled", text="🔄 Checking...")
        threading.Thread(target=self._update_check_worker, daemon=True).start()

    def _update_check_worker(self):
        try:
            info = updater.check_for_update(APP_VERSION)
            self.after(0, self._update_check_done, info, None)
        except updater.UpdateError as e:
            self.after(0, self._update_check_done, None, str(e))
        except Exception as e:
            self.after(0, self._update_check_done, None, str(e))

    def _update_check_done(self, info, error):
        self.update_btn.configure(state="normal", text="🔄 Check for Updates")

        if error:
            messagebox.showerror("Update Check Failed", error)
            return

        if not info["available"]:
            messagebox.showinfo(
                "You're up to date",
                f"You already have the latest version (v{APP_VERSION})."
            )
            return

        proceed = messagebox.askyesno(
            "Update Available",
            f"A new version is available: {info['latest_version']}\n"
            f"(you have v{APP_VERSION})\n\n"
            f"{info['notes'] or ''}\n\n"
            f"Download and install now? The app will restart automatically."
        )
        if not proceed:
            return

        if not updater.is_frozen():
            messagebox.showwarning(
                "Not Available",
                "Self-update only works in the built .exe — you're running from source (python app.py)."
            )
            return

        self._start_update_download(info["download_url"])

    def _start_update_download(self, download_url):
        self.update_btn.configure(state="disabled", text="⬇ Downloading 0%")
        threading.Thread(target=self._update_download_worker, args=(download_url,), daemon=True).start()

    def _update_download_worker(self, download_url):
        try:
            exe_dir = os.path.dirname(sys.executable)
            new_exe_path = os.path.join(exe_dir, "_update_download.exe")

            def progress_cb(pct, downloaded, total):
                self.after(0, lambda: self.update_btn.configure(text=f"⬇ Downloading {int(pct)}%"))

            updater.download_update(download_url, new_exe_path, progress_callback=progress_cb)
            self.after(0, self._update_download_done, new_exe_path, None)
        except Exception as e:
            self.after(0, self._update_download_done, None, str(e))

    def _update_download_done(self, new_exe_path, error):
        if error:
            self.update_btn.configure(state="normal", text="🔄 Check for Updates")
            messagebox.showerror("Update Failed", f"Could not download the update:\n\n{error}")
            return

        messagebox.showinfo(
            "Restarting...",
            "Update downloaded. The app will now close and restart with the new version."
        )
        try:
            updater.apply_update_and_restart(new_exe_path)
        except updater.UpdateError as e:
            self.update_btn.configure(state="normal", text="🔄 Check for Updates")
            messagebox.showerror("Update Failed", str(e))


# ---------------------------------------------------------------------------
# Entry Point
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    app = DocTypeReportApp()
    app.mainloop()
