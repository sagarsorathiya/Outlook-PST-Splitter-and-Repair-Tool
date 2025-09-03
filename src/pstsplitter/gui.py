"""Tkinter GUI for PST Splitter.

This file became corrupted in a prior styling refactor (mis-indented widget
construction at module scope). It has been repaired: all widget creation now
occurs inside _build_ui, state variables are initialized in __init__, and the
theme / preference loading sequence is restored.
"""
from __future__ import annotations

import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, font as tkfont
from pathlib import Path
import logging
import os
import sys
from typing import cast

from .splitter import split_pst, SplitResult, check_pst_health
from .util import configure_logging, LOG_QUEUE, load_prefs, save_prefs
from .outlook import is_outlook_available


class PSTSplitterApp(ttk.Frame):
    def __init__(self, master: tk.Tk):  # noqa: D401
        super().__init__(master, padding=8)
        root_win = cast(tk.Tk, self.master)
        root_win.title("PST Splitter - Ensue - Where Ideas Become Results")
        root_win.geometry("1200x600")
        root_win.minsize(1100, 550)  # Very compact for 15-inch screens
        
        # Professional styling
        root_win.configure(bg="#f0f2f5")
        self.grid(sticky="nsew")
        master.columnconfigure(0, weight=1)
        master.rowconfigure(0, weight=1)
        
        # Configure main grid for responsive layout
        self.columnconfigure(1, weight=2)  # Center panel gets more space
        self.columnconfigure((0,2), weight=1)
        self.rowconfigure(0, weight=1)

        # Core state variables
        self.source_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.csv_summary_var = tk.StringVar()
        self.size_var = tk.IntVar(value=500)  # default size
        self.size_unit_var = tk.StringVar(value="MB")
        self.mode_var = tk.StringVar(value="size")  # size|year|month|folder
        self.include_non_mail_var = tk.BooleanVar(value=False)
        self.move_items_var = tk.BooleanVar(value=False)
        self.verify_var = tk.BooleanVar(value=True)
        self.fast_enum_var = tk.BooleanVar(value=False)
        self.quiet_logs_var = tk.BooleanVar(value=False)
        self.stream_size_var = tk.BooleanVar(value=False)
        self.turbo_mode_var = tk.BooleanVar(value=False)  # New turbo mode option
        self.throttle_var = tk.IntVar(value=250)
        self.include_folders_var = tk.StringVar()
        self.exclude_folders_var = tk.StringVar()
        self.sender_domains_var = tk.StringVar()
        self.date_start_var = tk.StringVar()
        self.date_end_var = tk.StringVar()
        self.progress = tk.DoubleVar(value=0.0)

        # Runtime / UI state
        self.status_var = tk.StringVar(value="Ready - Select PST file to begin")
        self._throughput_var = tk.StringVar(value="")
        self.elapsed_eta_var = tk.StringVar(value="")
        self.filter_warn_var = tk.StringVar(value="")
        self.filters_summary_var = tk.StringVar(value="No active filters")
        self.estimate_var = tk.StringVar(value="")
        self.progress_stats_var = tk.StringVar(value="")
        self.find_var = tk.StringVar()
        self._sash_pos = None  # type: int | None
        # Dark theme removed (always light for visibility)
        self.dark_mode = False
        self.log_font_size = 10
        self._recent_sources = []  # type: list[str]
        # Persisted last folder include/exclude selections
        self._last_folder_include: set[str] = set()
        self._last_folder_exclude: set[str] = set()
        # Modern UI state variables
        self.csv_enabled = tk.BooleanVar(value=False)
        self._advanced_open = tk.BooleanVar(value=False)
        self.autoscroll_var = tk.BooleanVar(value=True)  # For log auto-scroll

        # Worker thread control
        self._cancel_event = threading.Event()
        self._worker = None  # type: threading.Thread | None
        self._start_time = None  # type: float | None
        self._is_cancelling = False
        self._last_cancel_request = 0.0

        # Build UI, then load preferences (which also applies theme & updates widgets)
        self._build_ui()
        self._load_preferences()
        self._update_estimate()
        self._apply_saved_geometry()
        
        # Adjust UI for screen size (important for 15-inch screens)
        self.after_idle(self._adjust_log_height_for_screen)

    # --- UI Construction -----------------------------------------------------------
    def _build_ui(self) -> None:
        # Helper: enhanced tooltip with better styling
        class Tooltip:
            def __init__(self, widget: tk.Widget, text: str):
                self.widget = widget
                self.text = text
                self.tip: tk.Toplevel | None = None
                widget.bind("<Enter>", self._show)
                widget.bind("<Leave>", self._hide)
            def _show(self, _e=None):
                if self.tip or not self.text:
                    return
                x = self.widget.winfo_rootx() + 10
                y = self.widget.winfo_rooty() + self.widget.winfo_height() + 6
                self.tip = tk.Toplevel(self.widget)
                self.tip.wm_overrideredirect(True)
                self.tip.wm_geometry(f"+{x}+{y}")
                lbl = tk.Label(self.tip, text=self.text, background="#2c3e50", foreground="white", 
                             relief="solid", borderwidth=0, padx=8, pady=4, font=("Segoe UI", 9))
                lbl.pack()
            def _hide(self, _e=None):
                if self.tip:
                    self.tip.destroy()
                    self.tip = None

        # Modern menu bar
        menu = tk.Menu(self.master)
        file_m = tk.Menu(menu, tearoff=0)
        file_m.add_command(label="📁 Open PST...", command=self._choose_source, accelerator="Ctrl+O")
        file_m.add_separator()
        file_m.add_command(label="❌ Exit", command=self.master.destroy)
        
        edit_m = tk.Menu(menu, tearoff=0)
        edit_m.add_command(label="🧹 Clear Log", command=self._clear_log, accelerator="Ctrl+L")
        edit_m.add_separator()
        edit_m.add_command(label="🔍 Zoom In", command=lambda: self.adjust_log_font(1))
        edit_m.add_command(label="🔍 Zoom Out", command=lambda: self.adjust_log_font(-1))
        
        help_m = tk.Menu(menu, tearoff=0)
        help_m.add_command(label="ℹ️ About", command=lambda: messagebox.showinfo("PST Splitter", 
            "PST Splitter v1.0\n"
            "Advanced Outlook PST Management Tool\n\n"
            "🚀 Features:\n"
            "• High-performance batch processing\n"
            "• Comprehensive analysis logging\n" 
            "• Enhanced year grouping\n"
            "• Export analysis reports\n\n"
            "👨‍💻 Developed by: Sagar Sorathiya\n"
            "🏢 Company: Ensue - Where Ideas Become Results\n\n"
            "© 2025 Sagar Sorathiya, Ensue. All rights reserved.\n"
            "Licensed under MIT License"))
        
        menu.add_cascade(label="File", menu=file_m)
        menu.add_cascade(label="Edit", menu=edit_m)
        menu.add_cascade(label="Help", menu=help_m)
        try:
            self.master.config(menu=menu)  # type: ignore[attr-defined]
        except Exception:
            pass

        # Enhanced keyboard shortcuts
        root = cast(tk.Tk, self.master)
        root.bind('<Control-o>', lambda e: self._choose_source())
        root.bind('<Control-r>', lambda e: self.start_split())
        root.bind('<Control-d>', lambda e: self.start_split(dry_run=True))
        root.bind('<Control-l>', lambda e: self._clear_log())
        root.bind('<F5>', lambda e: self.start_split())
        root.bind('<Escape>', lambda e: self.cancel_split())

        # Window close confirmation
        root.protocol("WM_DELETE_WINDOW", self._on_window_close)
        
        # Handle window resize to maintain log visibility on small screens
        root.bind("<Configure>", self._on_window_resize)

        # Main layout: 3-column design
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)
        self.columnconfigure(1, weight=2)
        self.columnconfigure(2, weight=1)

        # LEFT PANEL: Input & Configuration ================================
        left_panel = ttk.Frame(self, padding=(8, 8))
        left_panel.grid(row=0, column=0, sticky="nsew", padx=(0, 8))

        # File selection card with better spacing
        file_card = ttk.LabelFrame(left_panel, text=" 📁 Source & Output ", padding=(16, 12))
        file_card.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        file_card.columnconfigure(1, weight=1)

        ttk.Label(file_card, text="PST File:", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky="e", padx=(0,12), pady=(0,8))
        source_frame = ttk.Frame(file_card)
        source_frame.grid(row=0, column=1, columnspan=2, sticky="ew", pady=(0,8))
        source_frame.columnconfigure(0, weight=1)
        self.source_combo = ttk.Combobox(source_frame, textvariable=self.source_var, font=("Segoe UI", 10))
        self.source_combo.grid(row=0, column=0, sticky="ew", padx=(0,8))
        src_btn = ttk.Button(source_frame, text="Browse", command=self._choose_source, width=8)
        src_btn.grid(row=0, column=1)
        Tooltip(src_btn, "Select PST file to split")

        ttk.Label(file_card, text="Output:", font=("Segoe UI", 9, "bold")).grid(row=1, column=0, sticky="e", padx=(0,12), pady=(0,8))
        output_frame = ttk.Frame(file_card)
        output_frame.grid(row=1, column=1, columnspan=2, sticky="ew", pady=(0,8))
        output_frame.columnconfigure(0, weight=1)
        ttk.Entry(output_frame, textvariable=self.output_var, font=("Segoe UI", 10)).grid(row=0, column=0, sticky="ew", padx=(0,8))
        out_btn = ttk.Button(output_frame, text="Browse", command=self._choose_output, width=8)
        out_btn.grid(row=0, column=1)
        Tooltip(out_btn, "Choose destination folder")

        # Quick CSV toggle with better spacing
        csv_frame = ttk.Frame(file_card)
        csv_frame.grid(row=2, column=1, columnspan=2, sticky="w", pady=(6,0))
        ttk.Checkbutton(csv_frame, text="Generate summary CSV", variable=self.csv_enabled, 
                       command=self._toggle_csv).pack(side=tk.LEFT)

        # Split mode card with enhanced spacing
        mode_card = ttk.LabelFrame(left_panel, text=" ⚙️ Split Configuration ", padding=(16, 12))
        mode_card.grid(row=1, column=0, sticky="ew", pady=(0, 12))

        ttk.Label(mode_card, text="Split by:", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky="w", pady=(0,8))
        mode_frame = ttk.Frame(mode_card)
        mode_frame.grid(row=1, column=0, sticky="ew", pady=(0,12))
        mode_frame.columnconfigure((0,1), weight=1)
        
        # Mode buttons in 2x2 grid with better spacing
        modes = [("size", "📦 Size"), ("year", "📅 Year"), ("month", "📆 Month"), ("folder", "📂 Folder")]
        for i, (value, label) in enumerate(modes):
            row, col = divmod(i, 2)
            btn = ttk.Radiobutton(mode_frame, text=label, value=value, variable=self.mode_var, 
                                 command=self._toggle_mode_fields, style="Card.TRadiobutton")
            btn.grid(row=row, column=col, sticky="ew", padx=(0, 6 if col == 0 else 0), pady=(0, 4))

        # Size configuration with improved spacing
        size_frame = ttk.Frame(mode_card)
        size_frame.grid(row=2, column=0, sticky="ew", pady=(0,0))
        size_frame.columnconfigure(0, weight=1)
        ttk.Label(size_frame, text="Maximum size:", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky="w", pady=(0,6))
        size_input = ttk.Frame(size_frame)
        size_input.grid(row=1, column=0, sticky="w", pady=(0,0))
        self.size_entry = ttk.Entry(size_input, textvariable=self.size_var, width=8, font=("Segoe UI", 10, "bold"))
        self.size_entry.pack(side=tk.LEFT, padx=(0,8))
        self.size_unit_combo = ttk.Combobox(size_input, values=["MB", "GB", "TB"], textvariable=self.size_unit_var, 
                                           width=6, state="readonly", font=("Segoe UI", 10))
        self.size_unit_combo.pack(side=tk.LEFT)
        self._size_widgets = [self.size_entry, self.size_unit_combo]

        # Quick filters card with enhanced spacing
        filter_card = ttk.LabelFrame(left_panel, text=" 🔍 Filters ", padding=(16, 12))
        filter_card.grid(row=2, column=0, sticky="ew", pady=(0, 12))
        filter_card.columnconfigure(1, weight=1)

        # Essential options row with better spacing
        options_frame = ttk.Frame(filter_card)
        options_frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0,10))
        ttk.Checkbutton(options_frame, text="📧 Non-mail items", variable=self.include_non_mail_var).pack(side=tk.LEFT, padx=(0,16))
        ttk.Checkbutton(options_frame, text="📤 Move items", variable=self.move_items_var).pack(side=tk.LEFT)

        # Folder filters
        ttk.Label(filter_card, text="Folders:", font=("Segoe UI", 9)).grid(row=1, column=0, sticky="w", pady=2)
        folder_frame = ttk.Frame(filter_card)
        folder_frame.grid(row=1, column=1, sticky="ew", pady=2)
        folder_frame.columnconfigure(0, weight=1)
        ttk.Entry(folder_frame, textvariable=self.include_folders_var).grid(row=0, column=0, sticky="ew", padx=(0,4))
        ttk.Button(folder_frame, text="📋", command=self._fetch_and_select_folders, width=3).grid(row=0, column=1)

        # Enhanced date range with date pickers
        ttk.Label(filter_card, text="Date range:", font=("Segoe UI", 9)).grid(row=2, column=0, sticky="w", pady=2)
        date_frame = ttk.Frame(filter_card)
        date_frame.grid(row=2, column=1, sticky="ew", pady=2)
        
        # Start date
        start_frame = ttk.Frame(date_frame)
        start_frame.pack(side=tk.LEFT, padx=(0,8))
        ttk.Label(start_frame, text="From:", font=("Segoe UI", 8)).pack(side=tk.LEFT, padx=(0,2))
        self.date_start_entry = ttk.Entry(start_frame, textvariable=self.date_start_var, width=10)
        self.date_start_entry.pack(side=tk.LEFT)
        ttk.Button(start_frame, text="📅", command=lambda: self._pick_date("start"), width=3).pack(side=tk.LEFT, padx=(2,0))
        
        # End date  
        end_frame = ttk.Frame(date_frame)
        end_frame.pack(side=tk.LEFT)
        ttk.Label(end_frame, text="To:", font=("Segoe UI", 8)).pack(side=tk.LEFT, padx=(0,2))
        self.date_end_entry = ttk.Entry(end_frame, textvariable=self.date_end_var, width=10)
        self.date_end_entry.pack(side=tk.LEFT)
        ttk.Button(end_frame, text="📅", command=lambda: self._pick_date("end"), width=3).pack(side=tk.LEFT, padx=(2,0))
        
        # Quick date presets
        preset_frame = ttk.Frame(filter_card)
        preset_frame.grid(row=3, column=1, sticky="w", pady=(2,0))
        
        ttk.Button(preset_frame, text="This Year", command=lambda: self._set_date_preset("this_year"), style="Small.TButton").pack(side=tk.LEFT, padx=(0,4))
        ttk.Button(preset_frame, text="Last Year", command=lambda: self._set_date_preset("last_year"), style="Small.TButton").pack(side=tk.LEFT, padx=(0,4))
        ttk.Button(preset_frame, text="Last 30 Days", command=lambda: self._set_date_preset("last_30"), style="Small.TButton").pack(side=tk.LEFT, padx=(0,4))
        ttk.Button(preset_frame, text="Clear", command=lambda: self._set_date_preset("clear"), style="Small.TButton").pack(side=tk.LEFT)

        # Advanced filters (collapsible)
        adv_header = ttk.Frame(filter_card)
        adv_header.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(8,0))
        adv_toggle = ttk.Checkbutton(adv_header, text="⚡ Advanced options", variable=self._advanced_open, 
                                    command=self._toggle_advanced)
        adv_toggle.pack(side=tk.LEFT)

        self.advanced_frame = ttk.Frame(filter_card)
        af = self.advanced_frame  # shorthand
        af.columnconfigure(1, weight=1)
        
        # Folder exclusions
        ttk.Label(af, text="Exclude folders:", font=("Segoe UI", 9)).grid(row=0, column=0, sticky="w", pady=2)
        exclude_entry = ttk.Entry(af, textvariable=self.exclude_folders_var)
        exclude_entry.grid(row=0, column=1, sticky="ew", pady=2, padx=(4,0))
        Tooltip(exclude_entry, "Comma-separated folder names to exclude (e.g., 'Deleted Items, Junk Email')")
        
        # Sender domain filtering
        ttk.Label(af, text="Filter domains:", font=("Segoe UI", 9)).grid(row=1, column=0, sticky="w", pady=2)
        domain_entry = ttk.Entry(af, textvariable=self.sender_domains_var)
        domain_entry.grid(row=1, column=1, sticky="ew", pady=2, padx=(4,0))
        Tooltip(domain_entry, "Comma-separated domains to include (e.g., 'company.com, partner.org')")
        
        # Performance toggles with better organization
        perf_frame = ttk.LabelFrame(af, text=" Performance Options ", padding=(8, 6))
        perf_frame.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(8,0))
        
        perf_left = ttk.Frame(perf_frame)
        perf_left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        perf_right = ttk.Frame(perf_frame)
        perf_right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        fast_cb = ttk.Checkbutton(perf_left, text="🚀 Fast mode", variable=self.fast_enum_var)
        fast_cb.pack(anchor="w", pady=1)
        Tooltip(fast_cb, "Skip detailed enumeration for faster processing")
        
        quiet_cb = ttk.Checkbutton(perf_left, text="🔇 Quiet mode", variable=self.quiet_logs_var)
        quiet_cb.pack(anchor="w", pady=1)
        Tooltip(quiet_cb, "Reduce logging output for cleaner display")
        
        stream_cb = ttk.Checkbutton(perf_right, text="📊 Stream mode", variable=self.stream_size_var)
        stream_cb.pack(anchor="w", pady=1)
        Tooltip(stream_cb, "Optimize for very large PST files")
        
        turbo_cb = ttk.Checkbutton(perf_right, text="🏎️ Turbo mode", variable=self.turbo_mode_var)
        turbo_cb.pack(anchor="w", pady=1)
        Tooltip(turbo_cb, "Maximum performance mode (experimental)")

        # Validation warnings
        self.warning_label = ttk.Label(filter_card, textvariable=self.filter_warn_var, foreground="#e74c3c", font=("Segoe UI", 8))
        self.warning_label.grid(row=6, column=0, columnspan=2, sticky="w", pady=(4,0))

        # CENTER PANEL: Progress & Status ===================================
        center_panel = ttk.Frame(self, padding=(8, 8))
        center_panel.grid(row=0, column=1, sticky="nsew", padx=4)
        center_panel.rowconfigure(2, weight=1)  # Activity log expands but not too much

        # Large progress display with better spacing
        progress_card = ttk.LabelFrame(center_panel, text=" 🚀 Operation Status ", padding=(12, 8))
        progress_card.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        progress_card.columnconfigure(0, weight=1)

        # Status text with better spacing
        self.status_display = ttk.Label(progress_card, textvariable=self.status_var, 
                                       font=("Segoe UI", 14, "bold"), foreground="#2c3e50")
        self.status_display.grid(row=0, column=0, pady=(0,12))

        # Large progress bar with enhanced styling
        self.pb = ttk.Progressbar(progress_card, variable=self.progress, maximum=100, length=500, style="Large.Horizontal.TProgressbar")
        self.pb.grid(row=1, column=0, sticky="ew", pady=(0,12))

        # Progress stats with better spacing
        stats_frame = ttk.Frame(progress_card)
        stats_frame.grid(row=2, column=0, sticky="ew")
        stats_frame.columnconfigure((0,1,2), weight=1)
        
        self.throughput_label = ttk.Label(stats_frame, textvariable=self._throughput_var, font=("Segoe UI", 9), foreground="#7f8c8d")
        self.throughput_label.grid(row=0, column=0)
        
        self.eta_label = ttk.Label(stats_frame, textvariable=self.elapsed_eta_var, font=("Segoe UI", 9), foreground="#7f8c8d")
        self.eta_label.grid(row=0, column=1)
        
        self.estimate_label = ttk.Label(stats_frame, textvariable=self.estimate_var, font=("Segoe UI", 9), foreground="#95a5a6")
        self.estimate_label.grid(row=0, column=2)

        # Activity Log with enhanced spacing (moved from left panel for better visibility)
        log_card = ttk.LabelFrame(center_panel, text=" 📋 Activity Log ", padding=(12, 8))
        log_card.grid(row=2, column=0, sticky="nsew", pady=(0, 8))
        log_card.rowconfigure(1, weight=1)
        log_card.columnconfigure(0, weight=1)

        log_header = ttk.Frame(log_card)
        log_header.grid(row=0, column=0, sticky="ew", pady=(0,8))
        ttk.Label(log_header, text="Recent activity:", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT)
        ttk.Button(log_header, text="🧹 Clear", command=self._clear_log, width=8).pack(side=tk.RIGHT, padx=(4,0))
        ttk.Checkbutton(log_header, text="Auto-scroll", variable=self.autoscroll_var).pack(side=tk.RIGHT, padx=(8,4))

        # Enhanced log display with better sizing for center panel
        log_container = ttk.Frame(log_card)
        log_container.grid(row=1, column=0, sticky="nsew")
        log_container.rowconfigure(0, weight=1)
        log_container.columnconfigure(0, weight=1)
        
        self.log_text = tk.Text(log_container, height=8, state="disabled", wrap="word", 
                               font=("Consolas", 9), bg="#f8f9fa", fg="#2c3e50", 
                               relief="solid", borderwidth=1, padx=8, pady=6)
        log_scroll = ttk.Scrollbar(log_container, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scroll.set)
        self.log_text.grid(row=0, column=0, sticky="nsew")
        log_scroll.grid(row=0, column=1, sticky="ns")

        # RIGHT PANEL: Quick Actions =======================================
        right_panel = ttk.Frame(self, padding=(8, 8))
        right_panel.grid(row=0, column=2, sticky="nsew", padx=(8, 0))

        # Action buttons card with better spacing
        action_card = ttk.LabelFrame(right_panel, text=" ⚡ Quick Actions ", padding=(16, 12))
        action_card.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        action_card.columnconfigure(0, weight=1)

        # Large action buttons with enhanced spacing
        self.start_btn = ttk.Button(action_card, text="🚀 Start Split", command=self.start_split, style="Success.TButton")
        self.start_btn.grid(row=0, column=0, sticky="ew", pady=(0,8))
        Tooltip(self.start_btn, "Begin PST splitting operation")

        self.dry_btn = ttk.Button(action_card, text="🧪 Test Run", command=lambda: self.start_split(dry_run=True), style="Info.TButton")
        self.dry_btn.grid(row=1, column=0, sticky="ew", pady=(0,8))
        Tooltip(self.dry_btn, "Simulate operation without creating files")

        self.cancel_btn = ttk.Button(action_card, text="⏹️ Cancel", command=self.cancel_split, state="disabled", style="Warning.TButton")
        self.cancel_btn.grid(row=2, column=0, sticky="ew", pady=(0,8))
        Tooltip(self.cancel_btn, "Stop current operation")

        self.export_btn = ttk.Button(action_card, text="📊 Export Analysis", command=self.export_analysis, style="Info.TButton")
        self.export_btn.grid(row=3, column=0, sticky="ew", pady=(0,8))
        Tooltip(self.export_btn, "Export detailed analysis logs and performance data")

        self.repair_btn = ttk.Button(action_card, text="🔧 Repair PST", command=self.repair_pst, style="Warning.TButton")
        self.repair_btn.grid(row=4, column=0, sticky="ew", pady=(0,0))
        Tooltip(self.repair_btn, "Scan and repair PST file corruption (uses SCANPST.EXE)")

        # Info card with enhanced spacing
        info_card = ttk.LabelFrame(right_panel, text=" ℹ️ Tips ", padding=(16, 12))
        info_card.grid(row=1, column=0, sticky="ew", pady=(0, 12))
        
        tips = [
            "• Use Ctrl+O to open PST",
            "• F5 starts operation",
            "• Esc cancels operation",
            "• Test runs show structure",
            "• Stream mode for large PSTs"
        ]
        
        for tip in tips:
            ttk.Label(info_card, text=tip, font=("Segoe UI", 8), foreground="#7f8c8d").pack(anchor="w", pady=1)

        # Results Preview (moved below Tips for better layout)
        results_card = ttk.LabelFrame(right_panel, text=" 📊 Results Preview ", padding=(12, 8))
        results_card.grid(row=2, column=0, sticky="nsew")
        results_card.rowconfigure(0, weight=1)
        results_card.columnconfigure(0, weight=1)
        
        # Configure right panel row weights - increased weight for bigger results preview
        right_panel.rowconfigure(2, weight=1)

        # Summary text area - increased height for better visibility
        self.summary_text = tk.Text(results_card, height=8, state="disabled", wrap="word", 
                                   font=("Segoe UI", 9), bg="#f8f9fa", relief="solid", borderwidth=1)
        summary_scroll = ttk.Scrollbar(results_card, orient="vertical", command=self.summary_text.yview)
        self.summary_text.configure(yscrollcommand=summary_scroll.set)
        self.summary_text.grid(row=0, column=0, sticky="nsew")
        summary_scroll.grid(row=0, column=1, sticky="ns")

        # Bind events
        self.source_var.trace_add('write', lambda *_: self._update_estimate())
        self.mode_var.trace_add('write', lambda *_: self._update_estimate())
        self.size_var.trace_add('write', lambda *_: self._update_estimate())
        self.size_unit_var.trace_add('write', lambda *_: self._update_estimate())
        
        # Validation bindings
        for var in [self.include_folders_var, self.exclude_folders_var, self.sender_domains_var, 
                   self.date_start_var, self.date_end_var]:
            var.trace_add('write', lambda *_: self._validate_filters())

        # Initialize UI state
        self._toggle_mode_fields()
        self._validate_filters()
        self._toggle_advanced()  # Start collapsed
        self.apply_theme()

    # --- Actions -------------------------------------------------------------------
    def _choose_source(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("Outlook PST", "*.pst")])
        if path:
            self.source_var.set(path)
            self._remember_recent_source(path)
            # Automatically run health check on new PST
            self._show_pst_health_status(Path(path))

    def _choose_output(self) -> None:
        path = filedialog.askdirectory()
        if path:
            self.output_var.set(path)

    def _choose_csv(self) -> None:
        path = filedialog.asksaveasfilename(defaultextension='.csv', filetypes=[('CSV','*.csv')])
        if path:
            self.csv_summary_var.set(path)

    def start_split(self, dry_run: bool = False) -> None:
        """Enhanced start operation with better validation and feedback."""
        # Prevent multiple simultaneous operations
        if self._worker and self._worker.is_alive():
            messagebox.showwarning("Operation Running", 
                                 "An operation is already in progress.\n\n"
                                 "Please wait for it to complete or cancel it first.")
            return
            
        # Enhanced validation with specific error messages
        source = Path(self.source_var.get().strip())
        outdir = Path(self.output_var.get().strip())
        
        if not self.source_var.get().strip():
            messagebox.showerror("Missing Input", "Please select a source PST file.")
            return
            
        if not source.exists():
            messagebox.showerror("File Not Found", 
                               f"Source PST file not found:\n{source}\n\n"
                               "Please check the file path and try again.")
            return
            
        if not self.output_var.get().strip():
            messagebox.showerror("Missing Output", "Please select an output directory.")
            return
            
        if not outdir.exists():
            if messagebox.askyesno("Create Directory", 
                                 f"Output directory does not exist:\n{outdir}\n\n"
                                 "Would you like to create it?"):
                try:
                    outdir.mkdir(parents=True, exist_ok=True)
                except Exception as e:
                    messagebox.showerror("Directory Error", 
                                       f"Failed to create output directory:\n{e}")
                    return
            else:
                return
                
        if not is_outlook_available():
            messagebox.showerror("Outlook Required", 
                               "Microsoft Outlook is required but not available.\n\n"
                               "Please ensure Outlook is installed and try again.")
            return

        # Calculate size parameters
        max_size = self.size_var.get()
        mode = self.mode_var.get()
        unit = self.size_unit_var.get()
        
        if mode == "size" and max_size <= 0:
            messagebox.showerror("Invalid Size", "Please enter a valid size greater than 0.")
            self.size_entry.focus()
            return
            
        size_multipliers = {"MB": 1024**2, "GB": 1024**3, "TB": 1024**4}
        max_size_bytes = max_size * size_multipliers.get(unit, 1024**2) if mode == "size" else None

        # Reset cancellation state
        self._is_cancelling = False
        self._cancel_event.clear()
        
        # Enhanced UI feedback
        operation_type = "Test run" if dry_run else "Split operation"
        self.progress.set(0)
        self.status_var.set(f"🚀 Starting {operation_type.lower()}...")
        self._throughput_var.set("")
        self.elapsed_eta_var.set("")
        
        # Update button states
        self.start_btn.configure(state="disabled")
        self.dry_btn.configure(state="disabled")
        self.cancel_btn.configure(state="normal")
        
        # Log operation start
        self._append_log(f"Starting {operation_type}: {source.name}")
        if dry_run:
            self._append_log("DRY RUN - No files will be created")
        self._append_log(f"Output directory: {outdir}")
        self._append_log(f"Mode: {mode}" + (f" (max {max_size} {unit})" if mode == "size" else ""))
        import time as _t
        self._start_time = _t.time()

        def work():
            _did_com = False
            try:
                # Ensure COM initialized on this worker thread (required for Outlook automation)
                try:
                    import pythoncom  # type: ignore
                    pythoncom.CoInitialize()
                    _did_com = True  # flag for uninit
                except Exception:
                    pass
                include_set = {s.strip() for s in self.include_folders_var.get().split(',') if s.strip()}
                exclude_set = {s.strip() for s in self.exclude_folders_var.get().split(',') if s.strip()}
                result = split_pst(
                    source,
                    outdir,
                    mode,
                    max_size_bytes,
                    self._cancel_event,
                    progress_cb=self._on_progress,
                    dry_run=dry_run,
                    include_non_mail=self.include_non_mail_var.get(),
                    move_items=self.move_items_var.get(),
                    verify=self.verify_var.get(),
                    fast_enumeration=self.fast_enum_var.get(),
                    turbo_mode=self.turbo_mode_var.get(),
                    suppress_item_logs=self.quiet_logs_var.get(),
                    stream_size_mode=self.stream_size_var.get(),
                    throttle_progress_ms=self.throttle_var.get(),
                    include_folders=include_set if include_set else None,
                    exclude_folders=exclude_set if exclude_set else None,
                    sender_domains=self._parse_domains(),
                    date_range=self._parse_date_range(),
                    summary_csv=Path(self.csv_summary_var.get()) if self.csv_summary_var.get().strip() else None,
                )
                self._after_complete(result)
            except KeyboardInterrupt:
                # Operation was cancelled - this is expected behavior
                logging.info("Operation cancelled by user")
                self.master.after(0, self._handle_cancellation)
            except Exception as e:  # pragma: no cover - interactive
                logging.exception("Split failed")
                self.master.after(0, lambda: messagebox.showerror("Error", str(e)))
                self.master.after(0, self._reset_buttons)
            finally:
                if _did_com:
                    try:
                        import pythoncom  # type: ignore
                        pythoncom.CoUninitialize()
                    except Exception:
                        pass

        self._worker = threading.Thread(target=work, daemon=True)
        self._worker.start()
        self._poll_progress()
        self._poll_logs()

    def _poll_progress(self) -> None:
        """Enhanced progress polling with better responsiveness."""
        if self._worker and self._worker.is_alive():
            # Continue polling while operation is active
            self.master.after(150, self._poll_progress)  # Faster polling for better UX
        else:
            # Operation completed or stopped
            if self._is_cancelling:
                self._is_cancelling = False
                self.status_var.set("Operation cancelled")
                self.cancel_btn.configure(text="⏹️ Cancel Operation", state="disabled")
                self._append_log("Operation was cancelled by user")
            self._reset_buttons()

    def _poll_logs(self) -> None:
        """Enhanced log polling with better performance."""
        from queue import Empty
        drained = False
        lines_processed = 0
        max_lines_per_poll = 50  # Prevent UI freezing with large log bursts
        
        while lines_processed < max_lines_per_poll:
            try:
                rec = LOG_QUEUE.get_nowait()
            except Empty:
                break
            drained = True
            lines_processed += 1
            self._append_log(rec.getMessage())
            
        # Auto-scroll if enabled and we processed logs
        if drained and self.autoscroll_var.get():
            self.log_text.see("end")
            
        # Continue polling if operation is active
        if self._worker and self._worker.is_alive():
            self.master.after(200, self._poll_logs)  # Slightly slower for logs

    def _append_log(self, line: str) -> None:
        """Enhanced log appending with better formatting and limits."""
        import time
        
        # Add timestamp for better tracking
        timestamp = time.strftime("%H:%M:%S")
        formatted_line = f"[{timestamp}] {line}"
        
        self.log_text.configure(state="normal")
        
        # Limit log size to prevent memory issues
        line_count = int(self.log_text.index('end-1c').split('.')[0])
        if line_count > 1000:
            # Remove oldest 200 lines
            self.log_text.delete('1.0', '201.0')
            
        self.log_text.insert("end", formatted_line + "\n")
        self.log_text.configure(state="disabled")
        
        # Auto-scroll if enabled
        if self.autoscroll_var.get():
            self.log_text.see("end")

    def _after_complete(self, result: SplitResult) -> None:
        def update_ui():
            if self._cancel_event.is_set():
                self.status_var.set("Cancelled")
            else:
                summary = (
                    f"Done: {len(result.created_files)} parts, {result.total_items} items, "
                    f"{result.total_bytes/1024/1024:.1f}MB"
                )
                if result.errors:
                    summary += f", {len(result.errors)} errors"
                self.status_var.set(summary)
                if result.errors:
                    err_win = tk.Toplevel(self)
                    err_win.title("Copy Errors")
                    txt = tk.Text(err_win, width=100, height=20)
                    txt.pack(fill=tk.BOTH, expand=True)
                    txt.insert("end", "\n".join(result.errors))
                    txt.configure(state="disabled")
                # Populate summary tab
                try:
                    self._populate_summary(result)
                except Exception:
                    pass
            try:
                root = cast(tk.Tk, self.master)
                root.title("PST Splitter")
            except Exception:
                pass
            self.progress.set(100)
            self._reset_buttons()
        self.master.after(0, update_ui)

    # --- New UI Helper Methods ---------------------------------------------------
    def _toggle_csv(self) -> None:
        """Toggle CSV summary generation and auto-set filename."""
        if self.csv_enabled.get():
            if not self.csv_summary_var.get().strip():
                output_path = self.output_var.get().strip()
                if output_path:
                    from pathlib import Path
                    csv_path = Path(output_path) / "pst_split_summary.csv"
                    self.csv_summary_var.set(str(csv_path))

    def _toggle_advanced(self) -> None:
        """Show/hide advanced filter options."""
        try:
            if self._advanced_open.get():
                # Show advanced options
                self.advanced_frame.grid(row=5, column=0, columnspan=2, sticky="ew", pady=(4,0))
                # Ensure the parent frame updates its layout
                self.advanced_frame.update_idletasks()
            else:
                # Hide advanced options
                self.advanced_frame.grid_remove()
        except Exception as e:
            # Fallback in case of any issues
            print(f"Advanced toggle error: {e}")
            if hasattr(self, 'advanced_frame'):
                self.advanced_frame.grid_remove()

    # --- Folder Fetch / Selection ---------------------------------------------
    def _fetch_and_select_folders(self) -> None:
        """Fetch folder hierarchy and present dual include/exclude checkbox lists.

        Improvements over initial version:
          - Recursive enumeration (full paths like Inbox/Sub)
          - Aggregated counts including all descendant items per folder
          - Dual panes: Include vs Exclude (mutually exclusive selection enforced)
          - Progress bar with percentage & current folder indicator
          - Remembers last selections via _last_folder_include/_last_folder_exclude
        """
        if self._worker and self._worker.is_alive():
            messagebox.showinfo("Busy", "Please wait until current operation finishes.")
            return
        pst_path = Path(self.source_var.get())
        if not pst_path.exists():
            messagebox.showerror("Missing", "Select a source PST first.")
            return

        dialog = tk.Toplevel(self)
        dialog.title("Fetch Folders")
        dialog.geometry("560x520")
        try:
            dialog.transient(self.master)  # type: ignore[arg-type]
        except Exception:
            pass
        dialog.grab_set()
        ttk.Label(dialog, text=f"Enumerating folders in:\n{pst_path}").pack(anchor="w", padx=8, pady=(8,4))
        prog_frame = ttk.Frame(dialog)
        prog_frame.pack(fill=tk.X, padx=8)
        prog = ttk.Progressbar(prog_frame, mode="determinate", maximum=100)
        prog.pack(side=tk.LEFT, fill=tk.X, expand=True)
        pct_var = tk.StringVar(value="0%")
        ttk.Label(prog_frame, textvariable=pct_var, width=6).pack(side=tk.LEFT, padx=(6,0))
        cur_var = tk.StringVar(value="")
        ttk.Label(dialog, textvariable=cur_var, foreground="#555").pack(anchor="w", padx=8, pady=(2,6))

        panes = ttk.Panedwindow(dialog, orient=tk.HORIZONTAL)
        panes.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0,4))
        include_frame = ttk.Labelframe(panes, text="Include")
        exclude_frame = ttk.Labelframe(panes, text="Exclude")
        panes.add(include_frame, weight=1)
        panes.add(exclude_frame, weight=1)

        def make_scrollable(parent: ttk.Widget):
            wrapper = ttk.Frame(parent)
            wrapper.pack(fill=tk.BOTH, expand=True)
            wrapper.rowconfigure(0, weight=1)
            wrapper.columnconfigure(0, weight=1)
            ys = ttk.Scrollbar(wrapper, orient="vertical")
            canvas = tk.Canvas(wrapper, highlightthickness=0, yscrollcommand=ys.set)
            ys.config(command=canvas.yview)
            canvas.grid(row=0, column=0, sticky="nsew")
            ys.grid(row=0, column=1, sticky="ns")
            inner = ttk.Frame(canvas)
            canvas.create_window((0,0), window=inner, anchor="nw")
            inner.bind("<Configure>", lambda _e=None, c=canvas: c.configure(scrollregion=c.bbox("all")))
            return inner

        include_inner = make_scrollable(include_frame)
        exclude_inner = make_scrollable(exclude_frame)
        inc_vars: list[tuple[str, tk.BooleanVar]] = []
        exc_vars: list[tuple[str, tk.BooleanVar]] = []

        def toggle_peer(name: str, source: str) -> None:
            if source == 'inc':
                for n, v in exc_vars:
                    if n == name and v.get():
                        v.set(False)
            else:
                for n, v in inc_vars:
                    if n == name and v.get():
                        v.set(False)

        def populate(folder_list: list[tuple[str,int]]) -> None:
            for path, count in folder_list:
                inc_default = path in self._last_folder_include
                exc_default = (not inc_default) and path in self._last_folder_exclude
                iv = tk.BooleanVar(value=inc_default)
                ev = tk.BooleanVar(value=exc_default)
                ttk.Checkbutton(include_inner, text=f"{path} ({count})", variable=iv, command=lambda n=path: toggle_peer(n,'inc')).pack(anchor="w", padx=4, pady=1)
                ttk.Checkbutton(exclude_inner, text=f"{path} ({count})", variable=ev, command=lambda n=path: toggle_peer(n,'exc')).pack(anchor="w", padx=4, pady=1)
                inc_vars.append((path, iv))
                exc_vars.append((path, ev))

        # Worker for enumeration
        def worker() -> None:
            gathered: list[tuple[str,str,int]] = []  # (path, parent, direct_count)
            try:
                import pythoncom  # type: ignore
                pythoncom.CoInitialize()
                from win32com.client import Dispatch  # type: ignore
                app = Dispatch("Outlook.Application")
                ns = app.GetNamespace("MAPI")
                ns.AddStore(str(pst_path))
                target_store = None
                stores = ns.Stores
                for i in range(1, stores.Count + 1):  # type: ignore[attr-defined]
                    st = stores.Item(i)
                    try:
                        if Path(st.FilePath).resolve() == pst_path.resolve():  # type: ignore[attr-defined]
                            target_store = st
                            break
                    except Exception:
                        continue
                if not target_store:
                    raise RuntimeError("Could not locate attached store")
                root = target_store.GetRootFolder()
                # Pre-count folders for progress estimation
                try:
                    queue: list[tuple[object,str]] = [(root, "")]
                    total = 0
                    while queue:
                        f, rel = queue.pop(0)
                        total += 1
                        try:
                            sc = f.Folders.Count  # type: ignore[attr-defined]
                        except Exception:
                            sc = 0
                        for j in range(1, sc+1):
                            sub = f.Folders.Item(j)  # type: ignore[attr-defined]
                            relp = f"{rel}/{sub.Name}" if rel else sub.Name
                            queue.append((sub, relp))
                except Exception:
                    total = 1
                stack: list[tuple[object,str]] = [(root, "")]
                processed = 0
                while stack:
                    f, rel_path = stack.pop()
                    processed += 1
                    pct = int(processed / max(1,total) * 100)
                    self.master.after(0, lambda p=pct: prog.configure(value=p))
                    self.master.after(0, lambda p=pct: pct_var.set(f"{p}%"))
                    self.master.after(0, lambda r=rel_path or '/': cur_var.set(r[:60]))
                    try:
                        direct_items = f.Items.Count  # type: ignore[attr-defined]
                    except Exception:
                        direct_items = 0
                    parent_path = rel_path.rsplit('/',1)[0] if '/' in rel_path else ''
                    gathered.append((rel_path or '/', parent_path, int(direct_items)))
                    try:
                        sub_count = f.Folders.Count  # type: ignore[attr-defined]
                    except Exception:
                        sub_count = 0
                    for j in range(1, sub_count+1):
                        sub = f.Folders.Item(j)  # type: ignore[attr-defined]
                        sub_rel = f"{rel_path}/{sub.Name}" if rel_path else sub.Name
                        stack.append((sub, sub_rel))
                # Aggregate recursive counts
                from collections import defaultdict as _dd
                children: dict[str,list[str]] = _dd(list)
                direct: dict[str,int] = {}
                for path, parent, dc in gathered:
                    direct[path] = dc
                    children[parent].append(path)
                cache: dict[str,int] = {}
                def total_count(p: str) -> int:
                    if p in cache:
                        return cache[p]
                    t = direct.get(p,0)
                    for ch in children.get(p, []):
                        if ch != p:
                            t += total_count(ch)
                    cache[p] = t
                    return t
                result = [(p, total_count(p)) for p, _parent, _dc in gathered]
                result.sort(key=lambda x: x[0].lower())
                self.master.after(0, lambda r=result: populate(r))
            except Exception as e:  # pragma: no cover
                logging.exception("Folder fetch failed")
                self.master.after(0, lambda: messagebox.showerror("Fetch Failed", str(e)))
            finally:
                try:
                    ns.RemoveStore(root)  # type: ignore[name-defined]
                except Exception:
                    pass
                try:
                    import pythoncom  # type: ignore
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

        threading.Thread(target=worker, daemon=True).start()

        def apply_selection() -> None:
            inc_chosen = [n for n, v in inc_vars if v.get()]
            exc_chosen = [n for n, v in exc_vars if v.get()]
            self.include_folders_var.set(','.join(inc_chosen))
            self.exclude_folders_var.set(','.join(exc_chosen))
            self._last_folder_include = set(inc_chosen)
            self._last_folder_exclude = set(exc_chosen)
            dialog.destroy()
            self._validate_filters()
            self._save_preferences()

        btns = ttk.Frame(dialog)
        btns.pack(fill=tk.X, padx=8, pady=8)
        ttk.Button(btns, text="All Include", command=lambda: [v.set(True) for _n,v in inc_vars]).pack(side=tk.LEFT)
        ttk.Button(btns, text="Clear Include", command=lambda: [v.set(False) for _n,v in inc_vars]).pack(side=tk.LEFT, padx=(4,0))
        ttk.Button(btns, text="All Exclude", command=lambda: [v.set(True) for _n,v in exc_vars]).pack(side=tk.LEFT, padx=(12,0))
        ttk.Button(btns, text="Clear Exclude", command=lambda: [v.set(False) for _n,v in exc_vars]).pack(side=tk.LEFT, padx=(4,0))
        ttk.Button(btns, text="Apply", command=apply_selection).pack(side=tk.RIGHT)
        ttk.Button(btns, text="Cancel", command=dialog.destroy).pack(side=tk.RIGHT, padx=(0,6))

    def _populate_summary(self, result: SplitResult) -> None:
        if not hasattr(self, 'summary_text'):
            return
        lines = []
        lines.append("PST Split Summary")
        lines.append("=" * 60)
        lines.append(f"Total items processed: {result.total_items}")
        lines.append(f"Total bytes processed: {result.total_bytes} ({result.total_bytes/1024/1024:.2f} MB)")
        lines.append(f"Parts created: {len(result.created_files)}")
        for i, path in enumerate(result.created_files, 1):
            try:
                size = Path(path).stat().st_size
                lines.append(f"  {i}. {Path(path).name} - {size/1024/1024:.2f} MB")
            except Exception:
                lines.append(f"  {i}. {Path(path).name}")
        if result.errors:
            lines.append("")
            lines.append(f"Errors ({len(result.errors)}):")
            for e in result.errors[:50]:
                lines.append(f"  - {e}")
            if len(result.errors) > 50:
                lines.append(f"  ... {len(result.errors)-50} more")
        self.summary_text.configure(state='normal')
        self.summary_text.delete('1.0', 'end')
        self.summary_text.insert('end', "\n".join(lines))
        self.summary_text.configure(state='disabled')

    def _on_window_resize(self, event):
        """Handle window resize events to maintain proper layout on 15-inch screens."""
        # Only respond to window resize events, not child widget resizes
        if event.widget == self.master:
            # Get current window size
            window_height = self.master.winfo_height()
            window_width = self.master.winfo_width()
            
            # Adjust layout based on window size (helps with 15-inch screens)
            if window_height < 750:  # Small window height (typical 15-inch)
                # Make Activity Log more compact
                if hasattr(self, 'log_text'):
                    self.log_text.configure(height=10)
                # Make Results Preview more compact  
                if hasattr(self, 'summary_text'):
                    self.summary_text.configure(height=8)
            else:
                # Delay the adjustment to avoid too frequent calls
                self.after_idle(self._adjust_log_height_for_screen)

    def _adjust_log_height_for_screen(self):
        """Adjust log height based on screen size for better visibility."""
        try:
            screen_height = self.winfo_screenheight()
            
            # Calculate appropriate log height based on screen size
            # Very aggressive heights for 15-inch screen compatibility
            if screen_height <= 768:  # Typical 15-inch laptop resolution
                log_height = 6
                results_height = 4
            elif screen_height <= 900:
                log_height = 8
                results_height = 5
            elif screen_height <= 1080:
                log_height = 10
                results_height = 6
            else:
                log_height = 12
                results_height = 8
                
            # Update log text height if it exists
            if hasattr(self, 'log_text'):
                self.log_text.configure(height=log_height)
                
            # Update results text height if it exists
            if hasattr(self, 'summary_text'):
                self.summary_text.configure(height=results_height)
                
        except Exception:
            # Fallback to default height
            if hasattr(self, 'log_text'):
                self.log_text.configure(height=8)
            if hasattr(self, 'summary_text'):
                self.summary_text.configure(height=5)
                self.log_text.configure(height=15)

    def _pick_date(self, date_type):
        """Open a simple date picker dialog."""
        from tkinter import simpledialog
        from datetime import datetime
        
        title = "Select Start Date" if date_type == "start" else "Select End Date"
        current_val = self.date_start_var.get() if date_type == "start" else self.date_end_var.get()
        
        # Parse current date if valid
        current_date = None
        if current_val:
            try:
                current_date = datetime.strptime(current_val, "%Y-%m-%d")
            except ValueError:
                pass
        
        # Simple input dialog for now (can be enhanced with proper date picker)
        date_str = simpledialog.askstring(
            title, 
            f"Enter date (YYYY-MM-DD):\n\nCurrent: {current_val or 'None'}",
            initialvalue=current_val
        )
        
        if date_str:
            # Validate date format
            try:
                datetime.strptime(date_str, "%Y-%m-%d")
                if date_type == "start":
                    self.date_start_var.set(date_str)
                else:
                    self.date_end_var.set(date_str)
            except ValueError:
                messagebox.showerror("Invalid Date", "Please enter date in YYYY-MM-DD format")
    
    def _set_date_preset(self, preset):
        """Set date range based on preset selection."""
        from datetime import datetime, timedelta
        
        today = datetime.now()
        
        if preset == "this_year":
            start_date = datetime(today.year, 1, 1)
            end_date = datetime(today.year, 12, 31)
            self.date_start_var.set(start_date.strftime("%Y-%m-%d"))
            self.date_end_var.set(end_date.strftime("%Y-%m-%d"))
            
        elif preset == "last_year":
            last_year = today.year - 1
            start_date = datetime(last_year, 1, 1)
            end_date = datetime(last_year, 12, 31)
            self.date_start_var.set(start_date.strftime("%Y-%m-%d"))
            self.date_end_var.set(end_date.strftime("%Y-%m-%d"))
            
        elif preset == "last_30":
            end_date = today
            start_date = today - timedelta(days=30)
            self.date_start_var.set(start_date.strftime("%Y-%m-%d"))
            self.date_end_var.set(end_date.strftime("%Y-%m-%d"))
            
        elif preset == "clear":
            self.date_start_var.set("")
            self.date_end_var.set("")

    def _update_live_preview(self, done_items: int, total_items: int, done_bytes: int, total_bytes: int) -> None:
        """Update live progress in the results preview."""
        if not hasattr(self, 'summary_text'):
            return
        
        lines = []
        lines.append("Operation in Progress...")
        lines.append("=" * 60)
        
        # Progress stats
        progress_pct = (done_items / total_items * 100) if total_items else 0
        lines.append(f"Progress: {progress_pct:.1f}% ({done_items:,} / {total_items:,} items)")
        
        if done_bytes and total_bytes:
            data_pct = (done_bytes / total_bytes * 100)
            lines.append(f"Data: {data_pct:.1f}% ({done_bytes/1024/1024:.1f} / {total_bytes/1024/1024:.1f} MB)")
        
        # Time and speed info
        if self._start_time:
            import time as _t
            elapsed = _t.time() - self._start_time
            if elapsed > 0 and done_items > 0:
                items_per_sec = done_items / elapsed
                if done_bytes:
                    mb_per_sec = (done_bytes / 1024 / 1024) / elapsed
                    lines.append(f"Speed: {items_per_sec:.1f} items/s, {mb_per_sec:.2f} MB/s")
                else:
                    lines.append(f"Speed: {items_per_sec:.1f} items/s")
                
                # ETA calculation - only if we have meaningful speed data
                if total_items > done_items and items_per_sec > 0:
                    remaining_items = total_items - done_items
                    eta_seconds = remaining_items / items_per_sec
                    
                    if eta_seconds < 60:
                        eta_str = f"{eta_seconds:.0f}s"
                    elif eta_seconds < 3600:
                        eta_mins = int(eta_seconds // 60)
                        eta_secs = int(eta_seconds % 60)
                        eta_str = f"{eta_mins}m {eta_secs}s"
                    else:
                        eta_hours = int(eta_seconds // 3600)
                        eta_mins = int((eta_seconds % 3600) // 60)
                        eta_str = f"{eta_hours}h {eta_mins}m"
                    
                    lines.append(f"Estimated time remaining: {eta_str}")
            elif elapsed > 0:
                lines.append("Calculating speed...")
            else:
                lines.append("Starting operation...")
        
        lines.append("")
        lines.append("📊 Click 'Export Analysis' after completion for detailed report")
        
        # Update the text widget
        self.summary_text.configure(state='normal')
        self.summary_text.delete('1.0', 'end')
        self.summary_text.insert('end', "\n".join(lines))
        self.summary_text.configure(state='disabled')

    def _reset_buttons(self) -> None:
        self.start_btn.configure(state="normal")
        self.dry_btn.configure(state="normal")
        self.cancel_btn.configure(state="disabled")
        self._save_preferences()

    def cancel_split(self) -> None:
        """Enhanced cancellation with confirmation and proper cleanup."""
        import time
        from tkinter import messagebox
        current_time = time.time()
        
        # Prevent rapid multiple cancellation requests
        if current_time - self._last_cancel_request < 0.5:
            return
        self._last_cancel_request = current_time
        
        if not (self._worker and self._worker.is_alive()):
            # No operation running
            self.status_var.set("No operation to cancel")
            return
            
        if self._is_cancelling:
            # Already cancelling - show force option
            self.status_var.set("Force stopping...")
            # More aggressive cancellation approach could go here
            return
        
        # Ask for confirmation before cancelling
        result = messagebox.askyesno(
            "Confirm Cancellation", 
            "Are you sure you want to cancel the current operation?\n\nThis will stop the split process and any partially created PST files will be left in place.",
            icon="warning"
        )
        
        if not result:
            # User chose not to cancel
            return
        
        # Set cancellation flag
        self._is_cancelling = True
        self._cancel_event.set()
        
        # Immediate UI feedback
        self.status_var.set("🛑 Cancelling operation...")
        self.progress.set(0)
        self._throughput_var.set("")
        self.elapsed_eta_var.set("")
        
        # Disable start buttons, enable only cancel
        self.start_btn.configure(state="disabled")
        self.dry_btn.configure(state="disabled")
        self.cancel_btn.configure(text="⏹️ Force Stop", state="normal")
        
        # Schedule cleanup check
        self.master.after(1000, self._check_cancellation_complete)
        
        self._append_log("User requested cancellation...")

    def _check_cancellation_complete(self) -> None:
        """Check if cancellation completed and restore UI."""
        if not (self._worker and self._worker.is_alive()):
            # Operation has stopped
            self._is_cancelling = False
            self.status_var.set("Operation cancelled")
            self.start_btn.configure(state="normal")
            self.dry_btn.configure(state="normal") 
            self.cancel_btn.configure(text="⏹️ Cancel Operation", state="disabled")
            self._append_log("Operation cancelled successfully")
        else:
            # Still running, check again
            self.master.after(1000, self._check_cancellation_complete)

    def _handle_cancellation(self) -> None:
        """Handle proper cancellation cleanup when operation was cancelled."""
        self._is_cancelling = False
        self.status_var.set("Operation cancelled")
        self.progress.set(0)
        self._throughput_var.set("")
        self.elapsed_eta_var.set("")
        
        # Restore button states
        self.start_btn.configure(state="normal")
        self.dry_btn.configure(state="normal")
        self.cancel_btn.configure(text="⏹️ Cancel Operation", state="disabled")
        
        self._append_log("Operation cancelled by user")
        self._save_preferences()

    def export_analysis(self) -> None:
        """Export detailed analysis of the PST splitting session."""
        from tkinter import filedialog, messagebox
        import os
        from datetime import datetime
        from pathlib import Path
        
        try:
            # Check if there's any log data to export
            log_content = self.log_text.get("1.0", "end-1c").strip()
            if not log_content:
                messagebox.showwarning(
                    "No Data",
                    "No analysis data available. Please run a PST split operation first."
                )
                return
            
            # Get current timestamp for default filename
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            default_name = f"PST_Splitter_Analysis_{timestamp}"
            
            # Ask user to choose export location and format
            filename = filedialog.asksaveasfilename(
                title="Export Analysis Report",
                defaultextension=".txt",
                initialfile=default_name,
                filetypes=[
                    ("Text files", "*.txt"),
                    ("JSON files", "*.json"),
                    ("CSV files", "*.csv"),
                    ("All files", "*.*")
                ]
            )
            
            if not filename:
                return  # User cancelled
            
            # Create log exporter instance if not exists
            if not hasattr(self, '_log_exporter'):
                from .log_exporter import PSTSplitterLogExporter
                self._log_exporter = PSTSplitterLogExporter()
            
            # For now, create a simple text export of the log content
            # TODO: Integrate with proper log exporter for structured data
            file_path = Path(filename)
            
            if file_path.suffix.lower() == '.txt':
                # Simple text export
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(f"PST Splitter Analysis Report\n")
                    f.write(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                    f.write("=" * 50 + "\n\n")
                    f.write("Activity Log:\n")
                    f.write("-" * 20 + "\n")
                    f.write(log_content)
                success = True
            else:
                # For JSON/CSV, use the log exporter if session data exists
                if hasattr(self, '_log_exporter') and self._log_exporter.session_data.get('groups_created'):
                    # Use structured export
                    success = bool(self._log_exporter.export_analysis_report(file_path.parent))
                else:
                    # Fallback to text export
                    with open(file_path, 'w', encoding='utf-8') as f:
                        f.write(f"PST Splitter Activity Log - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
                        f.write(log_content)
                    success = True
            
            if success:
                messagebox.showinfo(
                    "Export Successful",
                    f"Analysis report exported successfully to:\n{filename}\n\nThe report includes:\n• Activity logs\n• Session details\n• Operation history"
                )
                
                # Optionally open the file location
                if messagebox.askyesno(
                    "Open Location", 
                    "Would you like to open the folder containing the exported file?"
                ):
                    import subprocess
                    import platform
                    folder_path = os.path.dirname(filename)
                    
                    if platform.system() == "Windows":
                        subprocess.run(["explorer", folder_path])
                    elif platform.system() == "Darwin":  # macOS
                        subprocess.run(["open", folder_path])
                    else:  # Linux and others
                        subprocess.run(["xdg-open", folder_path])
            else:
                messagebox.showerror(
                    "Export Failed",
                    "Failed to export analysis report. Please check the file path and try again."
                )
                
        except Exception as e:
            messagebox.showerror(
                "Export Error",
                f"An error occurred while exporting analysis:\n{str(e)}"
            )
            self._append_log(f"Analysis export error: {str(e)}")

    def repair_pst(self) -> None:
        """Launch PST repair utility (SCANPST.EXE) for the selected PST file."""
        from tkinter import messagebox
        import subprocess
        import os
        from pathlib import Path
        
        try:
            # Get the current PST file path
            pst_path = self.source_var.get().strip()
            if not pst_path:
                messagebox.showwarning(
                    "No PST File Selected",
                    "Please select a PST file first before attempting repair."
                )
                return
            
            if not os.path.exists(pst_path):
                messagebox.showerror(
                    "File Not Found",
                    f"The PST file does not exist:\n{pst_path}"
                )
                return
            
            # Check if operation is running
            if hasattr(self, '_worker') and self._worker and self._worker.is_alive():
                messagebox.showwarning(
                    "Operation in Progress",
                    "Cannot start PST repair while a split operation is running.\nPlease wait for the current operation to complete or cancel it first."
                )
                return
            
            # Search for SCANPST.EXE in common locations
            possible_paths = [
                # Office 365/2021/2019/2016
                r"C:\Program Files\Microsoft Office\root\Office16\SCANPST.EXE",
                r"C:\Program Files (x86)\Microsoft Office\root\Office16\SCANPST.EXE",
                # Office 2013
                r"C:\Program Files\Microsoft Office\Office15\SCANPST.EXE", 
                r"C:\Program Files (x86)\Microsoft Office\Office15\SCANPST.EXE",
                # Office 2010
                r"C:\Program Files\Microsoft Office\Office14\SCANPST.EXE",
                r"C:\Program Files (x86)\Microsoft Office\Office14\SCANPST.EXE",
                # Office 2007
                r"C:\Program Files\Microsoft Office\Office12\SCANPST.EXE",
                r"C:\Program Files (x86)\Microsoft Office\Office12\SCANPST.EXE",
            ]
            
            scanpst_path = None
            for path in possible_paths:
                if os.path.exists(path):
                    scanpst_path = path
                    break
            
            if not scanpst_path:
                # Try to search in PATH
                try:
                    result = subprocess.run(['where', 'scanpst.exe'], 
                                          capture_output=True, text=True, timeout=5)
                    if result.returncode == 0 and result.stdout.strip():
                        scanpst_path = result.stdout.strip().split('\n')[0]
                except:
                    pass
            
            if not scanpst_path:
                messagebox.showerror(
                    "SCANPST.EXE Not Found",
                    "Could not locate SCANPST.EXE (Outlook's PST repair tool).\n\n"
                    "Please ensure Microsoft Outlook is installed on this system.\n"
                    "You can also manually run SCANPST.EXE from your Office installation folder."
                )
                return
            
            # Confirm repair operation
            response = messagebox.askyesnocancel(
                "PST Repair Confirmation",
                f"This will launch Microsoft's PST repair tool (SCANPST.EXE) for:\n\n"
                f"{pst_path}\n\n"
                f"⚠️ IMPORTANT WARNINGS:\n"
                f"• ALWAYS backup your PST file before repair\n"
                f"• Close Outlook completely before starting repair\n"
                f"• Repair process may take a long time for large files\n"
                f"• Some data loss is possible during repair\n\n"
                f"Do you want to continue?\n\n"
                f"Choose:\n"
                f"• YES: Launch SCANPST.EXE now\n"
                f"• NO: Cancel repair\n"
                f"• CANCEL: Get backup instructions first"
            )
            
            if response is None:  # Cancel clicked - show backup instructions
                messagebox.showinfo(
                    "PST Backup Instructions",
                    "RECOMMENDED: Create a backup before repair\n\n"
                    "1. Close Outlook completely\n"
                    "2. Copy your PST file to a safe location:\n"
                    f"   {pst_path}\n"
                    "3. After backup, run PST repair again\n\n"
                    "This protects your data in case of repair issues."
                )
                return
            elif not response:  # No clicked
                return
            
            # Launch SCANPST.EXE
            self._append_log(f"Launching PST repair tool for: {os.path.basename(pst_path)}")
            
            # Launch with the PST file as parameter
            subprocess.Popen([scanpst_path, pst_path])
            
            messagebox.showinfo(
                "PST Repair Launched",
                f"Microsoft PST Repair Tool (SCANPST.EXE) has been launched.\n\n"
                f"File: {os.path.basename(pst_path)}\n\n"
                f"The repair tool will open in a separate window.\n"
                f"Follow the on-screen instructions to scan and repair your PST file.\n\n"
                f"⚠️ Remember: Keep Outlook closed during the repair process!"
            )
            
            self._append_log("PST repair tool launched successfully")
            
        except subprocess.TimeoutExpired:
            messagebox.showerror(
                "Timeout Error",
                "Timeout while searching for SCANPST.EXE. Please try again."
            )
            self._append_log("PST repair launch timeout")
        except Exception as e:
            messagebox.showerror(
                "Repair Error",
                f"Failed to launch PST repair tool:\n{str(e)}\n\n"
                f"You can manually run SCANPST.EXE from your Office installation folder."
            )
            self._append_log(f"PST repair error: {str(e)}")

    def _on_window_close(self) -> None:
        """Handle window close event with confirmation."""
        from tkinter import messagebox
        
        # Check if an operation is running - enhanced detection with multiple methods
        operation_running = False
        
        # Method 1: Check worker thread
        if hasattr(self, '_worker') and self._worker is not None and self._worker.is_alive():
            operation_running = True
        
        # Method 2: Check if cancel button is enabled (indicates operation is running)
        if hasattr(self, 'cancel_btn') and str(self.cancel_btn['state']) == 'normal':
            operation_running = True
        
        # Method 3: Check if start button is disabled (indicates operation is running)
        if hasattr(self, 'start_btn') and str(self.start_btn['state']) == 'disabled':
            operation_running = True
        
        if operation_running:
            result = messagebox.askyesno(
                "Confirm Exit", 
                "A PST split operation is currently running.\n\nAre you sure you want to exit?\n\nThis will cancel the operation and any partially created PST files will be left in place.",
                icon="warning"
            )
            
            if not result:
                # User chose not to exit
                return
            
            # Cancel the operation first
            if hasattr(self, '_cancel_event'):
                self._is_cancelling = True
                self._cancel_event.set()
            
            # Log the forced exit
            if hasattr(self, '_append_log'):
                self._append_log("Application closing - operation cancelled by user")
        else:
            # No operation running, ask for simple confirmation
            result = messagebox.askyesno(
                "Confirm Exit", 
                "Are you sure you want to exit PST Splitter?",
                icon="question"
            )
            
            if not result:
                # User chose not to exit
                return
        
        # Save preferences before closing
        if hasattr(self, '_save_preferences'):
            self._save_preferences()
        
        # Close the application
        self.master.destroy()

    def _toggle_mode_fields(self) -> None:
        # Disable size limit inputs when mode is not size
        is_size = self.mode_var.get() == "size"
        state = "normal" if is_size else "disabled"
        # Explicit references for size controls
        for wid in getattr(self, "_size_widgets", []):
            try:
                wid.configure(state=state)
            except Exception:
                pass
        self._update_estimate()

    def _update_estimate(self) -> None:
        """Update estimated number of output PST files (size mode only)."""
        try:
            if self.mode_var.get() != 'size':
                self.estimate_var.set('')
                return
            src = Path(self.source_var.get())
            if not src.is_file():
                self.estimate_var.set('')
                return
            size = src.stat().st_size
            try:
                val = int(self.size_var.get())
            except Exception:
                self.estimate_var.set('')
                return
            unit = self.size_unit_var.get()
            mul = 1024*1024
            if unit == 'GB':
                mul *= 1024
            elif unit == 'TB':
                mul *= 1024*1024
            if val <= 0:
                self.estimate_var.set('')
                return
            limit = val * mul
            import math
            est = max(1, math.ceil(size / limit))
            self.estimate_var.set(f"Est parts: {est}")
        except Exception:
            self.estimate_var.set('')

    # --- Progress & Prefs -------------------------------------------------------
    def _on_progress(self, done_items: int, total_items: int, done_bytes: int, total_bytes: int) -> None:
        pct = (done_items / total_items * 100.0) if total_items else 0.0
        self.master.after(0, lambda: self.progress.set(pct))
        if total_items:
            status = f"{done_items}/{total_items} items"
            if total_bytes:
                status += f" ({done_bytes/1024/1024:.1f}MB/{total_bytes/1024/1024:.1f}MB)"
            self.master.after(0, lambda s=status: self.status_var.set(s))
            try:
                root = cast(tk.Tk, self.master)
                root.after(0, lambda r=root, p=pct: r.title(f"PST Splitter - {p:.1f}%"))
            except Exception:
                pass
        
        # Update live preview in results area
        self.master.after(0, lambda: self._update_live_preview(done_items, total_items, done_bytes, total_bytes))
        
        # Throughput calculation
        if self._start_time and done_items:
            import time as _t
            elapsed = _t.time() - self._start_time
            if elapsed > 0:
                ips = done_items / elapsed
                mbps = (done_bytes / 1024 / 1024) / elapsed if done_bytes else 0
                tp = f"Throughput: {ips:.1f} items/s, {mbps:.2f} MB/s"
                self.master.after(0, lambda t=tp: self._throughput_var.set(t))
        # Elapsed / ETA
        if self._start_time:
            import time as _t
            el = _t.time() - self._start_time
            
            # Format time in minutes and seconds
            def format_time(seconds):
                if seconds < 60:
                    return f"{seconds:.0f}s"
                elif seconds < 3600:
                    mins = int(seconds // 60)
                    secs = int(seconds % 60)
                    return f"{mins}m {secs}s"
                else:
                    hours = int(seconds // 3600)
                    mins = int((seconds % 3600) // 60)
                    return f"{hours}h {mins}m"
            
            if total_items and done_items:
                est_total = el * (total_items / (done_items or 1))
                remain = max(0.0, est_total - el)
                
                el_str = format_time(el)
                eta_str = format_time(remain)
                msg = f"Elapsed: {el_str}  ETA: {eta_str}"
                stats = f"{(done_items/el):.1f} items/s | ETA {eta_str}"
            else:
                el_str = format_time(el) if el > 0 else "0s"
                msg = f"Elapsed: {el_str}"
                stats = f"{(done_items/el):.1f} items/s" if el > 0 and done_items else ""
            self.master.after(0, lambda m=msg: self.elapsed_eta_var.set(m))
            self.master.after(0, lambda s=stats: self.progress_stats_var.set(s))

    def _load_preferences(self) -> None:
        data = load_prefs()
        self.source_var.set(data.get("source", ""))
        self.output_var.set(data.get("output", ""))
        self.size_var.set(int(data.get("size", 500)))
        self.size_unit_var.set(data.get("unit", "MB"))
        self.mode_var.set(data.get("mode", "size"))
        self.include_non_mail_var.set(bool(data.get("incl_non_mail", False)))
        self.move_items_var.set(bool(data.get("move_items", False)))
        self.verify_var.set(data.get("verify", True))
        self.fast_enum_var.set(data.get("fast_enum", False))
        self.turbo_mode_var.set(data.get("turbo_mode", False))
        self.quiet_logs_var.set(data.get("quiet_logs", False))
        self.stream_size_var.set(data.get("stream_size", False))
        self.throttle_var.set(int(data.get("throttle_ms", 250)))
        self.include_folders_var.set(data.get("include_folders", ""))
        self.exclude_folders_var.set(data.get("exclude_folders", ""))
        # Restore previously selected folder sets if present
        self._last_folder_include = set(data.get("_last_folder_include", [])) if isinstance(data.get("_last_folder_include", []), list) else set()
        self._last_folder_exclude = set(data.get("_last_folder_exclude", [])) if isinstance(data.get("_last_folder_exclude", []), list) else set()
        self.sender_domains_var.set(data.get("sender_domains", ""))
        self.date_start_var.set(data.get("date_start", ""))
        self.date_end_var.set(data.get("date_end", ""))
        self.csv_summary_var.set(data.get("csv_summary", ""))
        self._recent_sources = []
        rs = data.get("recent_sources", [])
        if isinstance(rs, list):
            self._recent_sources = [p for p in rs if isinstance(p, str)][:10]
        if hasattr(self, 'source_combo'):
            try:
                self.source_combo.configure(values=self._recent_sources)
            except Exception:
                pass
        self.dark_mode = bool(data.get("dark_mode", False))
        self.log_font_size = int(data.get("log_font_size", 10))
        # Apply preferences to UI elements
        self._apply_log_font()
        self._toggle_mode_fields()
        self.apply_theme()
        self._update_filters_summary()

    def _save_preferences(self) -> None:
        save_prefs(
            {
                "source": self.source_var.get(),
                "output": self.output_var.get(),
                "size": self.size_var.get(),
                "unit": self.size_unit_var.get(),
                "mode": self.mode_var.get(),
                "incl_non_mail": self.include_non_mail_var.get(),
                "move_items": self.move_items_var.get(),
                "verify": self.verify_var.get(),
                "fast_enum": self.fast_enum_var.get(),
                "turbo_mode": self.turbo_mode_var.get(),
                "quiet_logs": self.quiet_logs_var.get(),
                "stream_size": self.stream_size_var.get(),
                "throttle_ms": self.throttle_var.get(),
                "include_folders": self.include_folders_var.get(),
                "exclude_folders": self.exclude_folders_var.get(),
                "_last_folder_include": sorted(getattr(self, '_last_folder_include', set())),
                "_last_folder_exclude": sorted(getattr(self, '_last_folder_exclude', set())),
                "sender_domains": self.sender_domains_var.get(),
                "date_start": self.date_start_var.get(),
                "date_end": self.date_end_var.get(),
                "csv_summary": self.csv_summary_var.get(),
                "recent_sources": getattr(self, '_recent_sources', []),
                "dark_mode": int(self.dark_mode),
                "log_font_size": self.log_font_size,
            }
        )

    def _clear_log(self) -> None:
        """Clear the log display with user confirmation for active operations."""
        # Don't clear during active operations unless user confirms
        if self._worker and self._worker.is_alive():
            if not messagebox.askyesno("Clear Log", 
                                     "Operation is running. Clear log anyway?\n\n"
                                     "This will remove current progress information."):
                return
                
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")
        
        # Add a cleared indicator
        self._append_log("Activity log cleared by user")

    # --- New helpers ----------------------------------------------------------
    def _validate_filters(self) -> None:
        inc = {s.strip().lower() for s in self.include_folders_var.get().split(',') if s.strip()}
        exc = {s.strip().lower() for s in self.exclude_folders_var.get().split(',') if s.strip()}
        overlap = inc.intersection(exc)
        warn = []
        if overlap:
            warn.append(f"Overlap ignored: {', '.join(sorted(overlap))}")
        # Date validation
        for lbl, val in [("start", self.date_start_var.get()), ("end", self.date_end_var.get())]:
            if val.strip():
                if not self._is_valid_date(val.strip()):
                    warn.append(f"Invalid {lbl} date '{val}'")
        self.filter_warn_var.set("; ".join(warn))
        self._update_filters_summary()

    def toggle_theme(self) -> None:  # retained for backward compatibility, no-op
        return

    def apply_theme(self) -> None:
        style = ttk.Style()
        try:
            style.theme_use('clam')
        except Exception:
            pass
        
        # Modern color scheme
        bg = "#f8f9fa"
        card_bg = "#ffffff"
        text_bg = "#ffffff"
        fg = "#2c3e50"
        accent = "#3498db"
        success = "#27ae60"
        warning = "#e67e22"
        danger = "#e74c3c"
        muted = "#7f8c8d"
        
        # Configure modern styles
        style.configure('TFrame', background=bg)
        style.configure('TLabelframe', background=bg, foreground=fg)
        style.configure('TLabelframe.Label', background=bg, foreground=fg, font=("Segoe UI", 9, "bold"))
        style.configure('TLabel', background=bg, foreground=fg, font=("Segoe UI", 9))
        
        # Modern button styles
        style.configure('TButton', font=("Segoe UI", 9), padding=(8,4))
        style.configure('Success.TButton', background=success, foreground="white", font=("Segoe UI", 10, "bold"))
        style.map('Success.TButton', background=[('active', '#229954')])
        style.configure('Info.TButton', background=accent, foreground="white", font=("Segoe UI", 9))
        style.map('Info.TButton', background=[('active', '#2980b9')])
        style.configure('Warning.TButton', background=warning, foreground="white", font=("Segoe UI", 9))
        style.map('Warning.TButton', background=[('active', '#d35400')])
        style.configure('Small.TButton', font=("Segoe UI", 8), padding=(6,2))
        
        # Modern progress bar
        style.configure('Large.Horizontal.TProgressbar', background=accent, lightcolor=accent, darkcolor=accent)
        
        # Card-style radiobuttons and checkbuttons
        style.configure('Card.TRadiobutton', background=card_bg, foreground=fg, font=("Segoe UI", 9), padding=(8,4))
        style.map('Card.TRadiobutton', background=[('active', '#ecf0f1')])
        
        # Entry styling
        style.configure('TEntry', fieldbackground=text_bg, foreground=fg, font=("Segoe UI", 9), padding=(4,2))
        style.configure('TCombobox', fieldbackground=text_bg, foreground=fg, font=("Segoe UI", 9))
        
        # Apply to text widgets
        if hasattr(self, 'log_text'):
            self.log_text.configure(bg=text_bg, fg=fg, insertbackground=fg, font=("Consolas", 9))
        if hasattr(self, 'summary_text'):
            self.summary_text.configure(bg=text_bg, fg=fg, insertbackground=fg, font=("Segoe UI", 9))
        
        try:
            self.master.configure(bg=bg)  # type: ignore[attr-defined]
        except Exception:
            pass

    def adjust_log_font(self, delta: int) -> None:
        self.log_font_size = max(6, min(40, self.log_font_size + delta))
        self._apply_log_font()

    def _apply_log_font(self) -> None:
        if hasattr(self, 'log_text'):
            f = tkfont.Font(family='Consolas', size=self.log_font_size)
            self.log_text.configure(font=f)

    def _apply_saved_geometry(self) -> None:
        try:
            data = load_prefs()
            geom = data.get('geometry')
            if geom:
                root = cast(tk.Tk, self.master)
                root.geometry(geom)
        except Exception:
            pass

    # Recent sources & filters summary helpers
    def _remember_recent_source(self, path: str) -> None:
        try:
            if path in self._recent_sources:
                self._recent_sources.remove(path)
            self._recent_sources.insert(0, path)
            self._recent_sources = self._recent_sources[:10]
            if hasattr(self, 'source_combo'):
                self.source_combo.configure(values=self._recent_sources)
            self._save_preferences()
        except Exception:
            pass

    def _update_filters_summary(self) -> None:
        parts: list[str] = []
        if self.include_folders_var.get().strip():
            parts.append(f"inc:[{self.include_folders_var.get()}]")
        if self.exclude_folders_var.get().strip():
            parts.append(f"exc:[{self.exclude_folders_var.get()}]")
        if self.sender_domains_var.get().strip():
            parts.append(f"dom:[{self.sender_domains_var.get()}]")
        ds = self.date_start_var.get().strip()
        de = self.date_end_var.get().strip()
        if ds or de:
            parts.append(f"date:{ds or '...'}->{de or '...'}")
        summary = "No filters" if not parts else " | ".join(parts)
        if len(summary) > 120:
            summary = summary[:117] + "..."
        try:
            self.filters_summary_var.set(summary)
        except Exception:
            pass

    # --- Parsing helpers ---------------------------------------------------
    def _is_valid_date(self, txt: str) -> bool:
        if len(txt) != 10:
            return False
        from datetime import datetime as _dt
        try:
            _dt.strptime(txt, '%Y-%m-%d')
            return True
        except Exception:
            return False

    def _parse_domains(self) -> set[str] | None:
        raw = self.sender_domains_var.get().strip()
        if not raw:
            return None
        out = {s.strip().lstrip('@').lower() for s in raw.split(',') if s.strip()}
        return out or None

    def _parse_date_range(self):  # -> tuple[datetime|None, datetime|None] | None
        s = self.date_start_var.get().strip()
        e = self.date_end_var.get().strip()
        if not s and not e:
            return None
        from datetime import datetime as _dt
        start_dt = None
        end_dt = None
        if s:
            if not self._is_valid_date(s):
                raise ValueError(f"Invalid start date '{s}' (YYYY-MM-DD)")
            start_dt = _dt.strptime(s, '%Y-%m-%d')
        if e:
            if not self._is_valid_date(e):
                raise ValueError(f"Invalid end date '{e}' (YYYY-MM-DD)")
            end_dt = _dt.strptime(e, '%Y-%m-%d')
        if start_dt and end_dt and end_dt < start_dt:
            raise ValueError('End date is before start date')
        return (start_dt, end_dt)

    def _show_pst_health_status(self, pst_path: Path) -> None:
        """Display PST health check results to the user."""
        try:
            health_report = check_pst_health(pst_path)
            
            # Prepare health status message
            title = "PST Health Check"
            size_info = health_report.get("size_info", {})
            
            message_parts = []
            if size_info:
                message_parts.append(f"📊 PST Analysis:")
                message_parts.append(f"• Type: {size_info.get('pst_type', 'Unknown')}")
                message_parts.append(f"• Size: {size_info.get('current_size_formatted', 'Unknown')}")
                message_parts.append(f"• Free Space: {size_info.get('free_space_formatted', 'Unknown')}")
                message_parts.append(f"• Utilization: {size_info.get('utilization_percent', 0):.1f}%")
                message_parts.append("")
            
            # Add warnings if any
            warnings = health_report.get("warnings", [])
            if warnings:
                message_parts.append("⚠️ Warnings:")
                for warning in warnings:
                    message_parts.append(f"• {warning}")
                message_parts.append("")
            
            # Add recommendations if any
            recommendations = health_report.get("recommendations", [])
            if recommendations:
                message_parts.append("💡 Recommendations:")
                for rec in recommendations:
                    message_parts.append(f"• {rec}")
            
            message = "\n".join(message_parts)
            
            # Show appropriate message box based on health status
            if not health_report.get("healthy", True):
                messagebox.showwarning(title, f"❌ Health Check Issues Detected\n\n{message}")
            elif warnings:
                messagebox.showinfo(title, f"⚠️ Health Check - Cautions\n\n{message}")
            else:
                messagebox.showinfo(title, f"✅ PST Health Check Passed\n\n{message}")
                
        except Exception as e:
            logging.exception("Error during PST health check")
            messagebox.showerror("Health Check Error", f"Could not analyze PST health:\n{e}")


def run_app() -> None:
    # Force DEBUG logging by default (Feature A)
    import logging as _logging
    stop_event = configure_logging(level=_logging.DEBUG)
    root = tk.Tk()
    app = PSTSplitterApp(root)
    root.protocol("WM_DELETE_WINDOW", root.destroy)
    root.mainloop()
    stop_event.set()


if __name__ == "__main__":  # pragma: no cover
    run_app()
