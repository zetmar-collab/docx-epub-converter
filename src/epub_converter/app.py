"""Tkinter desktop application for DOCX to EPUB conversion."""

from __future__ import annotations

import html
import re
import threading
import webbrowser
from datetime import date
from pathlib import Path
from tkinter import filedialog, messagebox
import tkinter as tk
from tkinter import ttk

from PIL import Image, ImageTk

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    _HAS_DND = True
except ImportError:
    _HAS_DND = False

from epub_converter.config import load_config, update_config
from epub_converter.constants import (
    APP_TITLE,
    APP_VERSION,
    DOC_LANGUAGES,
    NAVY_ACCENT,
    NAVY_ACCENT_ACTIVE,
    NAVY_BG,
    NAVY_BORDER,
    NAVY_INPUT,
    NAVY_PANEL,
    NAVY_PANEL_ALT,
    NAVY_TEXT,
)
from epub_converter.docx_parser import parse_docx
from epub_converter.epub_builder import build_epub
from epub_converter.i18n import TRANSLATIONS
from epub_converter.isbn import is_valid_isbn, normalize_isbn
from epub_converter.models import ConversionResult
from epub_converter.preflight import inspect_docx_for_epub
from epub_converter.preview import write_preview_html
from epub_converter.ui_fonts import pick_ui_font
from epub_converter.utils import (
    default_output_dir,
    html_to_plain_text,
    open_path,
    safe_filename,
    sanitize_dropped_path,
)
from epub_converter.validation import epubcheck_available, validate_epub

_BaseApp = TkinterDnD.Tk if _HAS_DND else tk.Tk  # type: ignore[misc]


class EpubConverterApp(_BaseApp):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_TITLE} v{APP_VERSION}")
        self.geometry("1120x760")
        self.minsize(900, 620)

        self.ui_lang_var  = tk.StringVar(value="pl")
        self.doc_lang_var = tk.StringVar(value="pl")

        self.docx_path   = tk.StringVar()
        self.cover_path  = tk.StringVar()
        self.output_path = tk.StringVar(value=str(default_output_dir() / "ebook.epub"))
        self.title_var     = tk.StringVar()
        self.author_var    = tk.StringVar()
        self.publisher_var = tk.StringVar()
        self.year_var      = tk.StringVar(value=str(date.today().year))
        self.isbn_var      = tk.StringVar()
        self.status_var    = tk.StringVar(value=TRANSLATIONS["pl"]["status_ready"])

        self.cover_preview: ImageTk.PhotoImage | None = None
        self.result: ConversionResult | None = None
        self.preview_html_path: Path | None = None
        self.convert_buttons: list[ttk.Button] = []
        self._tw: dict[str, tk.Widget] = {}
        self._doc_lang_codes: list[str] = [code for _, code in DOC_LANGUAGES]
        self._output_manually_set = False
        self._java_available = epubcheck_available()
        self._ui_font = pick_ui_font(self, 10)
        self._ui_font_bold = pick_ui_font(self, 10, "bold")
        self._ui_font_title = pick_ui_font(self, 18, "bold")
        self._ui_font_section = pick_ui_font(self, 11, "bold")
        self._ui_font_small = pick_ui_font(self, 9, "bold")

        _cfg = load_config()
        self._last_docx_dir:  str = _cfg.get("last_docx_dir", "")
        self._last_cover_dir: str = _cfg.get("last_cover_dir", "")
        self._profile_doc_lang: str = _cfg.get("profile_doc_lang", "pl")

        if _cfg.get("profile_author"):
            self.author_var.set(_cfg["profile_author"])
        if _cfg.get("profile_publisher"):
            self.publisher_var.set(_cfg["profile_publisher"])
        if _cfg.get("profile_year"):
            self.year_var.set(_cfg["profile_year"])

        self.title_var.trace_add("write", self._on_title_change)

        self._configure_style()
        self._build_ui()
        if not self._java_available:
            messagebox.showwarning(self.t("java_warn_title"), self.t("java_warn_msg"), parent=self)

    def _save_profile(self) -> None:
        update_config({
            "profile_author":    self.author_var.get().strip(),
            "profile_publisher": self.publisher_var.get().strip(),
            "profile_year":      self.year_var.get().strip(),
            "profile_doc_lang":  self.doc_lang_var.get(),
        })

    # ------------------------------------------------------------------
    # Translations
    # ------------------------------------------------------------------

    def t(self, key: str) -> str:
        return TRANSLATIONS[self.ui_lang_var.get()].get(key, key)

    # ------------------------------------------------------------------
    # Styles
    # ------------------------------------------------------------------

    def _configure_style(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        self.configure(background=NAVY_BG)
        style.configure("TFrame", background=NAVY_BG)
        style.configure("Card.TFrame", background=NAVY_PANEL, relief="solid", borderwidth=1)
        style.configure("TLabel", background=NAVY_BG, foreground=NAVY_TEXT, font=self._ui_font)
        style.configure("Card.TLabel", background=NAVY_PANEL, foreground=NAVY_TEXT, font=self._ui_font)
        style.configure("Title.TLabel", background=NAVY_BG, font=self._ui_font_title, foreground=NAVY_TEXT)
        style.configure("Section.TLabel", background=NAVY_PANEL, font=self._ui_font_section, foreground=NAVY_TEXT)
        style.configure(
            "TEntry",
            fieldbackground=NAVY_INPUT,
            foreground=NAVY_BG,
            bordercolor=NAVY_BORDER,
            lightcolor=NAVY_BORDER,
            darkcolor=NAVY_BORDER,
            insertcolor=NAVY_BG,
        )
        style.configure(
            "TButton",
            background=NAVY_PANEL_ALT,
            foreground=NAVY_TEXT,
            bordercolor=NAVY_BORDER,
            focusthickness=1,
            focuscolor=NAVY_ACCENT,
            padding=(10, 5),
        )
        style.map(
            "TButton",
            background=[("active", NAVY_ACCENT), ("disabled", "#1b3048")],
            foreground=[("disabled", "#788da6")],
        )
        style.configure(
            "Primary.TButton",
            font=self._ui_font_bold,
            background=NAVY_ACCENT,
            foreground="#ffffff",
            bordercolor=NAVY_ACCENT_ACTIVE,
        )
        style.map("Primary.TButton", background=[("active", NAVY_ACCENT_ACTIVE), ("disabled", "#1b3048")])
        style.configure(
            "Lang.TButton",
            font=self._ui_font_small,
            background=NAVY_PANEL_ALT,
            foreground=NAVY_TEXT,
            bordercolor=NAVY_BORDER,
            padding=(6, 3),
        )
        style.map("Lang.TButton", background=[("active", NAVY_ACCENT_ACTIVE)])
        style.configure("Status.TLabel", background=NAVY_PANEL_ALT, foreground=NAVY_TEXT, padding=8)
        style.configure("TSeparator", background=NAVY_BORDER)
        style.configure("Vertical.TScrollbar", background=NAVY_PANEL_ALT, troughcolor=NAVY_BG, bordercolor=NAVY_BORDER)
        style.configure(
            "TCombobox",
            fieldbackground=NAVY_INPUT,
            background=NAVY_PANEL_ALT,
            foreground=NAVY_BG,
            selectbackground=NAVY_ACCENT,
            selectforeground="#ffffff",
            bordercolor=NAVY_BORDER,
        )
        style.map("TCombobox", fieldbackground=[("readonly", NAVY_INPUT)], foreground=[("readonly", NAVY_BG)])
        style.configure(
            "TProgressbar",
            troughcolor=NAVY_BG,
            background=NAVY_ACCENT,
            bordercolor=NAVY_BORDER,
            lightcolor=NAVY_ACCENT,
            darkcolor=NAVY_ACCENT,
        )

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    def _build_ui(self):
        canvas = tk.Canvas(self, bg=NAVY_BG, highlightthickness=0)
        vscroll = ttk.Scrollbar(self, orient=tk.VERTICAL, command=canvas.yview)
        canvas.configure(yscrollcommand=vscroll.set)
        vscroll.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        root = ttk.Frame(canvas, padding=16)
        _root_id = canvas.create_window((0, 0), window=root, anchor=tk.NW)

        def _on_frame_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _on_canvas_configure(event):
            canvas.itemconfig(_root_id, width=event.width)

        def _on_mousewheel(event):
            if not isinstance(event.widget, (tk.Text, tk.Listbox)):
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        def _bind_wheel(_event):
            canvas.bind_all("<MouseWheel>", _on_mousewheel)

        def _unbind_wheel(_event):
            canvas.unbind_all("<MouseWheel>")

        root.bind("<Configure>", _on_frame_configure)
        canvas.bind("<Configure>", _on_canvas_configure)
        canvas.bind("<Enter>", _bind_wheel)
        canvas.bind("<Leave>", _unbind_wheel)

        header = ttk.Frame(root)
        header.pack(fill=tk.X, pady=(0, 12))

        title_group = ttk.Frame(header)
        title_group.pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Label(title_group, text=f"{APP_TITLE} v{APP_VERSION}", style="Title.TLabel").pack(anchor=tk.W)
        self._tw["app_subtitle"] = ttk.Label(title_group, text=self.t("app_subtitle"))
        self._tw["app_subtitle"].pack(anchor=tk.W, pady=(4, 0))
        self._tw["app_author_label"] = ttk.Label(title_group, text=self.t("app_author_label"))
        self._tw["app_author_label"].pack(anchor=tk.W, pady=(4, 0))

        right_group = ttk.Frame(header)
        right_group.pack(side=tk.RIGHT)
        lang_btn = ttk.Button(right_group, text="EN", command=self._toggle_ui_lang, style="Lang.TButton", width=4)
        lang_btn.pack(anchor=tk.E, pady=(0, 6))
        self._tw["lang_toggle"] = lang_btn
        self._add_convert_button(right_group, anchor=tk.E, ipadx=18, ipady=6)

        main = ttk.PanedWindow(root, orient=tk.HORIZONTAL)
        main.pack(fill=tk.BOTH, expand=True)

        left  = ttk.Frame(main, style="Card.TFrame", padding=14)
        right = ttk.Frame(main, style="Card.TFrame", padding=14)
        main.add(left, weight=1)
        main.add(right, weight=2)

        self._build_form(left)
        self._build_preview(right)

        bottom = ttk.Frame(root)
        bottom.pack(fill=tk.X, pady=(12, 0))
        self._add_convert_button(bottom, side=tk.RIGHT, ipadx=24, ipady=7)

        self.progress = ttk.Progressbar(root, mode="indeterminate", length=100)
        self.progress.pack(fill=tk.X, pady=(8, 0))

        ttk.Label(root, textvariable=self.status_var, style="Status.TLabel").pack(fill=tk.X, pady=(4, 0))

    def _add_convert_button(self, parent, **pack_options):
        button = ttk.Button(
            parent,
            text=self.t("btn_convert"),
            command=self.convert,
            style="Primary.TButton",
        )
        self.convert_buttons.append(button)
        if pack_options:
            button.pack(**pack_options)
        return button

    def _build_form(self, parent):
        self._tw["section_metadata"] = ttk.Label(parent, text=self.t("section_metadata"), style="Section.TLabel")
        self._tw["section_metadata"].pack(anchor=tk.W)

        self._tw["field_title"]     = self._entry(parent, "field_title",     self.title_var)
        self._tw["field_author"]    = self._entry(parent, "field_author",    self.author_var)
        self._tw["field_publisher"] = self._entry(parent, "field_publisher", self.publisher_var)
        self._tw["field_year"]      = self._entry(parent, "field_year",      self.year_var)
        _isbn_hdr = ttk.Frame(parent, style="Card.TFrame")
        _isbn_hdr.pack(fill=tk.X, pady=(8, 2))
        self._tw["field_isbn"] = ttk.Label(_isbn_hdr, text=self.t("field_isbn"), style="Card.TLabel")
        self._tw["field_isbn"].pack(side=tk.LEFT)
        self._tw["isbn_help_btn"] = ttk.Button(
            _isbn_hdr, text="?", width=2, style="Lang.TButton", command=self._show_isbn_help
        )
        self._tw["isbn_help_btn"].pack(side=tk.RIGHT)
        ttk.Entry(parent, textvariable=self.isbn_var).pack(fill=tk.X)

        self._tw["field_desc"] = ttk.Label(parent, text=self.t("field_desc"), style="Card.TLabel")
        self._tw["field_desc"].pack(anchor=tk.W, pady=(8, 2))
        self.desc_text = self._create_text_widget(parent, height=5)
        self.desc_text.pack(fill=tk.X)

        ttk.Separator(parent).pack(fill=tk.X, pady=14)
        self._tw["section_lang_doc"] = ttk.Label(parent, text=self.t("section_lang_doc"), style="Section.TLabel")
        self._tw["section_lang_doc"].pack(anchor=tk.W)
        lang_display_values = [name for name, _ in DOC_LANGUAGES]
        self.doc_lang_combo = ttk.Combobox(parent, values=lang_display_values, state="readonly")
        saved_lang = getattr(self, "_profile_doc_lang", "pl")
        if saved_lang in self._doc_lang_codes:
            self.doc_lang_combo.current(self._doc_lang_codes.index(saved_lang))
            self.doc_lang_var.set(saved_lang)
        else:
            self.doc_lang_combo.current(0)
        self.doc_lang_combo.pack(fill=tk.X, pady=(4, 0))
        self.doc_lang_combo.bind("<<ComboboxSelected>>", self._on_doc_lang_change)

        ttk.Separator(parent).pack(fill=tk.X, pady=14)
        self._tw["section_files"] = ttk.Label(parent, text=self.t("section_files"), style="Section.TLabel")
        self._tw["section_files"].pack(anchor=tk.W)
        self._tw["field_docx"], self._tw["btn_docx"] = self._path_picker(
            parent, "field_docx", self.docx_path, self.pick_docx, dnd_handler=self._dnd_docx
        )
        self._tw["field_cover"], self._tw["btn_cover"] = self._path_picker(
            parent, "field_cover", self.cover_path, self.pick_cover, dnd_handler=self._dnd_cover
        )

        self.cover_label = ttk.Label(parent, text=self.t("no_cover_preview"), style="Card.TLabel")
        self.cover_label.pack(anchor=tk.W, pady=(8, 0))
        self.cover_warn_label = ttk.Label(parent, text="", style="Card.TLabel", foreground="#e07000")
        self.cover_warn_label.pack(anchor=tk.W)

        ttk.Separator(parent).pack(fill=tk.X, pady=14)
        self._tw["section_save"] = ttk.Label(parent, text=self.t("section_save"), style="Section.TLabel")
        self._tw["section_save"].pack(anchor=tk.W)
        self._tw["field_output"], self._tw["btn_output"] = self._path_picker(
            parent, "field_output", self.output_path, self.pick_output_path
        )

        form_button = self._add_convert_button(parent)
        form_button.pack(fill=tk.X, pady=(16, 2), ipady=7)
        self._tw["btn_batch"] = ttk.Button(parent, text=self.t("btn_batch"), command=self.open_batch_dialog)
        self._tw["btn_batch"].pack(fill=tk.X, pady=(0, 6))

        actions = ttk.Frame(parent, style="Card.TFrame")
        actions.pack(fill=tk.X, pady=(8, 0))
        self.open_file_button = ttk.Button(
            actions, text=self.t("btn_open_epub"), command=self.open_epub, state=tk.DISABLED
        )
        self.open_file_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
        self.open_folder_button = ttk.Button(
            actions, text=self.t("btn_open_folder"), command=self.open_output_folder, state=tk.DISABLED
        )
        self.open_folder_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(4, 0))

    def _build_preview(self, parent):
        self._tw["section_preview"] = ttk.Label(parent, text=self.t("section_preview"), style="Section.TLabel")
        self._tw["section_preview"].pack(anchor=tk.W)

        meta_frame = ttk.Frame(parent, style="Card.TFrame")
        meta_frame.pack(fill=tk.X, pady=(8, 10))
        self.result_label = ttk.Label(meta_frame, text=self.t("preview_after"), style="Card.TLabel")
        self.result_label.pack(anchor=tk.W)

        chapter_frame = ttk.Frame(parent, style="Card.TFrame")
        chapter_frame.pack(fill=tk.BOTH, expand=True)

        left = ttk.Frame(chapter_frame, style="Card.TFrame")
        left.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        self._tw["toc_label"] = ttk.Label(left, text=self.t("toc_label"), style="Section.TLabel")
        self._tw["toc_label"].pack(anchor=tk.W)
        self.chapter_list = tk.Listbox(
            left,
            width=28,
            activestyle="dotbox",
            exportselection=False,
            bg=NAVY_INPUT,
            fg=NAVY_BG,
            selectbackground=NAVY_ACCENT,
            selectforeground="#ffffff",
            highlightbackground=NAVY_BORDER,
            highlightcolor=NAVY_ACCENT,
            relief=tk.SOLID,
            borderwidth=1,
        )
        self.chapter_list.pack(fill=tk.BOTH, expand=True, pady=(6, 0))
        self.chapter_list.bind("<<ListboxSelect>>", self.on_chapter_select)

        text_frame = ttk.Frame(chapter_frame, style="Card.TFrame")
        text_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.preview_text = self._create_text_widget(text_frame)
        self.preview_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=self.preview_text.yview)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.preview_text.configure(yscrollcommand=scroll.set)
        self.preview_text.insert(tk.END, self.t("preview_placeholder"))
        self.preview_text.configure(state=tk.DISABLED)

        self.open_html_button = ttk.Button(
            parent, text=self.t("btn_html_preview"), command=self.open_html_preview, state=tk.DISABLED
        )
        self.open_html_button.pack(anchor=tk.E, pady=(10, 0))

    def _create_text_widget(self, parent, height=None):
        kwargs = {
            "wrap": tk.WORD,
            "relief": tk.SOLID,
            "borderwidth": 1,
            "bg": NAVY_INPUT,
            "fg": NAVY_BG,
            "insertbackground": NAVY_BG,
            "selectbackground": NAVY_ACCENT,
            "selectforeground": "#ffffff",
            "highlightbackground": NAVY_BORDER,
            "highlightcolor": NAVY_ACCENT,
            "font": self._ui_font,
        }
        if height is not None:
            kwargs["height"] = height
        return tk.Text(parent, **kwargs)

    def _entry(self, parent, label_key: str, variable) -> ttk.Label:
        lbl = ttk.Label(parent, text=self.t(label_key), style="Card.TLabel")
        lbl.pack(anchor=tk.W, pady=(8, 2))
        ttk.Entry(parent, textvariable=variable).pack(fill=tk.X)
        return lbl

    def _path_picker(
        self, parent, label_key: str, variable, command, dnd_handler=None
    ) -> tuple[ttk.Label, ttk.Button]:
        lbl = ttk.Label(parent, text=self.t(label_key), style="Card.TLabel")
        lbl.pack(anchor=tk.W, pady=(8, 2))
        row = ttk.Frame(parent, style="Card.TFrame")
        row.pack(fill=tk.X)
        entry = ttk.Entry(row, textvariable=variable)
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        if _HAS_DND and dnd_handler is not None:
            entry.drop_target_register(DND_FILES)  # type: ignore[attr-defined]
            entry.dnd_bind("<<Drop>>", dnd_handler)  # type: ignore[attr-defined]
        btn = ttk.Button(row, text=self.t("btn_pick"), command=command)
        btn.pack(side=tk.LEFT, padx=(8, 0))
        return lbl, btn

    # ------------------------------------------------------------------
    # Language switching
    # ------------------------------------------------------------------

    def _toggle_ui_lang(self):
        self.ui_lang_var.set("en" if self.ui_lang_var.get() == "pl" else "pl")
        self._apply_language()

    def _apply_language(self):
        lang = self.ui_lang_var.get()
        tr = TRANSLATIONS[lang]

        self._tw["lang_toggle"].configure(text="EN" if lang == "pl" else "PL")

        label_keys = [
            "app_subtitle", "app_author_label",
            "section_metadata", "field_title", "field_author", "field_publisher",
            "field_year", "field_isbn", "field_desc", "section_lang_doc",
            "section_files", "field_docx", "field_cover",
            "section_save", "field_output",
            "section_preview", "toc_label", "btn_batch",
        ]
        for key in label_keys:
            if key in self._tw:
                self._tw[key].configure(text=tr[key])

        for key in ("btn_docx", "btn_cover", "btn_output"):
            if key in self._tw:
                self._tw[key].configure(text=tr["btn_pick"])

        for btn in self.convert_buttons:
            btn.configure(text=tr["btn_convert"])

        self.open_file_button.configure(text=tr["btn_open_epub"])
        self.open_folder_button.configure(text=tr["btn_open_folder"])
        self.open_html_button.configure(text=tr["btn_html_preview"])

        if self.result is None:
            self.status_var.set(tr["status_ready"])
            self.result_label.configure(text=tr["preview_after"])
            self.preview_text.configure(state=tk.NORMAL)
            self.preview_text.delete("1.0", tk.END)
            self.preview_text.insert(tk.END, tr["preview_placeholder"])
            self.preview_text.configure(state=tk.DISABLED)

        if not self.cover_path.get():
            self.cover_label.configure(text=tr["no_cover_preview"])

    def _on_doc_lang_change(self, _event):
        self.doc_lang_var.set(self._doc_lang_codes[self.doc_lang_combo.current()])

    # ------------------------------------------------------------------
    # Auto output filename from title
    # ------------------------------------------------------------------

    def _show_isbn_help(self):
        self.show_issue_window(self.t("isbn_help_title"), self.t("isbn_help_text"))

    def _on_title_change(self, *_):
        if self._output_manually_set:
            return
        title = self.title_var.get().strip()
        if not title:
            return
        current = self.output_path.get()
        try:
            out_dir = Path(current).parent
        except Exception:
            out_dir = default_output_dir()
        self.output_path.set(str(out_dir / safe_filename(title)))

    # ------------------------------------------------------------------
    # File pickers & DnD
    # ------------------------------------------------------------------

    def pick_docx(self):
        init_dir = self._last_docx_dir or str(default_output_dir())
        path = filedialog.askopenfilename(
            title=self.t("dlg_pick_docx"),
            initialdir=init_dir,
            filetypes=[("DOCX", "*.docx")],
        )
        if path:
            self._set_docx(path)

    def _dnd_docx(self, event):
        path = sanitize_dropped_path(event.data)
        if path is None or path.suffix.lower() != ".docx":
            messagebox.showwarning(APP_TITLE, self.t("err_invalid_dnd"), parent=self)
            return
        self._set_docx(str(path))

    def _set_docx(self, path: str):
        self.docx_path.set(path)
        self._last_docx_dir = str(Path(path).parent)
        update_config({"last_docx_dir": self._last_docx_dir})

    def pick_cover(self):
        init_dir = self._last_cover_dir or self._last_docx_dir or str(default_output_dir())
        path = filedialog.askopenfilename(
            title=self.t("dlg_pick_cover"),
            initialdir=init_dir,
            filetypes=[("Images / Obrazy", "*.png *.jpg *.jpeg")],
        )
        if path:
            self._set_cover(path)

    def _dnd_cover(self, event):
        path = sanitize_dropped_path(event.data)
        if path is None or path.suffix.lower() not in {".png", ".jpg", ".jpeg"}:
            messagebox.showwarning(APP_TITLE, self.t("err_invalid_dnd"), parent=self)
            return
        self._set_cover(str(path))

    def _set_cover(self, path: str):
        self.cover_path.set(path)
        self._last_cover_dir = str(Path(path).parent)
        update_config({"last_cover_dir": self._last_cover_dir})
        self.load_cover_preview(Path(path))

    def pick_output_path(self):
        current = self.output_path.get()
        try:
            init_dir  = str(Path(current).parent) if current else str(default_output_dir())
            init_file = Path(current).name if current else "ebook.epub"
        except Exception:
            init_dir  = str(default_output_dir())
            init_file = "ebook.epub"
        path = filedialog.asksaveasfilename(
            title=self.t("dlg_pick_output"),
            initialdir=init_dir,
            initialfile=init_file,
            defaultextension=".epub",
            filetypes=[("EPUB", "*.epub")],
        )
        if path:
            self.output_path.set(path)
            self._output_manually_set = True

    def load_cover_preview(self, path: Path):
        try:
            image = Image.open(path)
            orig_w, orig_h = image.size
            image.thumbnail((160, 220))
            self.cover_preview = ImageTk.PhotoImage(image)
            self.cover_label.configure(image=self.cover_preview, text="")
            warnings = []
            if orig_h > 0 and abs(orig_w / orig_h - 2 / 3) > 0.10:
                warnings.append(self.t("cover_warn_ratio"))
            if orig_w < 1200:
                warnings.append(self.t("cover_warn_res"))
            self.cover_warn_label.configure(text="\n".join(warnings))
        except Exception as exc:
            self.cover_label.configure(image="", text=self.t("cover_load_error") + str(exc))
            self.cover_warn_label.configure(text="")

    # ------------------------------------------------------------------
    # Validation
    # ------------------------------------------------------------------

    def validate_input(self) -> bool:
        missing = []
        if not self.title_var.get().strip():
            missing.append(self.t("missing_title"))
        if not self.author_var.get().strip():
            missing.append(self.t("missing_author"))
        if not self.publisher_var.get().strip():
            missing.append(self.t("missing_publisher"))
        if not self.year_var.get().strip():
            missing.append(self.t("missing_year"))
        if not self.desc_text.get("1.0", tk.END).strip():
            missing.append(self.t("missing_desc"))
        if not Path(self.docx_path.get()).is_file():
            missing.append(self.t("missing_docx"))
        if not Path(self.cover_path.get()).is_file():
            missing.append(self.t("missing_cover"))
        if not self.output_path.get().strip():
            missing.append(self.t("missing_output"))
        if missing:
            self.show_issue_window(
                self.t("err_missing_title"),
                self.t("err_missing_msg") + "\n- ".join(missing),
            )
            return False

        if not re.fullmatch(r"\d{4}", self.year_var.get().strip()):
            self.show_issue_window(self.t("err_year_title"), self.t("err_year_msg"))
            return False

        isbn_raw = self.isbn_var.get().strip()
        if isbn_raw and not is_valid_isbn(isbn_raw):
            self.show_issue_window(self.t("err_isbn_title"), self.t("err_isbn_msg"))
            return False

        try:
            with Image.open(Path(self.cover_path.get())) as image:
                if image.format not in {"PNG", "JPEG"}:
                    self.show_issue_window(self.t("err_cover_format_title"), self.t("err_cover_format_msg"))
                    return False
        except Exception as exc:
            self.show_issue_window(self.t("err_cover_open_title"), self.t("err_cover_open_msg") + str(exc))
            return False

        return True

    def preflight_docx(self) -> bool:
        tr = TRANSLATIONS[self.ui_lang_var.get()]
        try:
            docx_bytes = Path(self.docx_path.get()).read_bytes()
            issues = inspect_docx_for_epub(docx_bytes, tr)
        except Exception as exc:
            self.show_issue_window(self.t("preflight_err_title"), self.t("preflight_err_msg") + str(exc))
            return False
        if issues:
            message = self.t("preflight_msg") + "\n\n".join(
                f"{i}. {issue}" for i, issue in enumerate(issues, 1)
            )
            return self.show_preflight_dialog(self.t("preflight_title"), message)
        return True

    def show_preflight_dialog(self, title: str, message: str) -> bool:
        """Return True if user chooses to convert despite warnings."""
        result: list[bool] = [False]

        window = tk.Toplevel(self)
        window.title(title)
        window.configure(background=NAVY_BG)
        window.geometry("620x420")
        window.minsize(520, 320)
        window.transient(self)
        window.grab_set()

        frame = ttk.Frame(window, padding=16)
        frame.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frame, text=title, style="Section.TLabel").pack(anchor=tk.W, pady=(0, 8))

        text_frame = ttk.Frame(frame)
        text_frame.pack(fill=tk.BOTH, expand=True)
        text = self._create_text_widget(text_frame)
        text.insert(tk.END, message)
        text.configure(state=tk.DISABLED)
        text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=text.yview)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        text.configure(yscrollcommand=scroll.set)

        btn_row = ttk.Frame(frame)
        btn_row.pack(anchor=tk.E, pady=(12, 0))

        def _copy_to_clipboard():
            window.clipboard_clear()
            window.clipboard_append(message)
            window.update()

        def _continue():
            result[0] = True
            window.destroy()

        ttk.Button(btn_row, text=self.t("btn_copy"), command=_copy_to_clipboard).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(btn_row, text=self.t("btn_preflight_cancel"), command=window.destroy).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(
            btn_row, text=self.t("btn_preflight_continue"), command=_continue, style="Primary.TButton"
        ).pack(side=tk.LEFT)
        window.wait_window()
        return result[0]

    def show_issue_window(self, title: str, message: str):
        window = tk.Toplevel(self)
        window.title(title)
        window.configure(background=NAVY_BG)
        window.geometry("620x420")
        window.minsize(520, 320)
        window.transient(self)
        window.grab_set()

        frame = ttk.Frame(window, padding=16)
        frame.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frame, text=title, style="Section.TLabel").pack(anchor=tk.W, pady=(0, 8))

        text_frame = ttk.Frame(frame)
        text_frame.pack(fill=tk.BOTH, expand=True)
        text = self._create_text_widget(text_frame)
        text.insert(tk.END, message)
        text.configure(state=tk.DISABLED)
        text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=text.yview)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        text.configure(yscrollcommand=scroll.set)

        btn_row = ttk.Frame(frame)
        btn_row.pack(anchor=tk.E, pady=(12, 0))

        def _copy_to_clipboard():
            window.clipboard_clear()
            window.clipboard_append(message)
            window.update()

        ttk.Button(btn_row, text=self.t("btn_copy"), command=_copy_to_clipboard).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(btn_row, text=self.t("btn_ok"), command=window.destroy).pack(side=tk.LEFT)
        window.focus_set()

    # ------------------------------------------------------------------
    # Conversion
    # ------------------------------------------------------------------

    def convert(self):
        if not self.validate_input():
            return
        if not self.preflight_docx():
            return
        self._set_convert_buttons_state(tk.DISABLED)
        self.status_var.set(self.t("converting"))
        self.result_label.configure(text=self.t("converting"))
        self.progress.start(10)
        threading.Thread(target=self._convert_worker, daemon=True).start()

    def _set_convert_buttons_state(self, state):
        for button in self.convert_buttons:
            button.configure(state=state)

    def _convert_worker(self):
        try:
            docx_path  = Path(self.docx_path.get())
            cover_path = Path(self.cover_path.get())
            doc_lang   = self.doc_lang_var.get()

            saved_path = Path(self.output_path.get()).expanduser()
            if not saved_path.suffix:
                saved_path = saved_path.with_suffix(".epub")
            saved_path.parent.mkdir(parents=True, exist_ok=True)

            meta = {
                "title":       self.title_var.get().strip(),
                "author":      self.author_var.get().strip(),
                "publisher":   self.publisher_var.get().strip() or self.author_var.get().strip(),
                "year":        self.year_var.get().strip() or str(date.today().year),
                "isbn":        normalize_isbn(self.isbn_var.get()),
                "description": self.desc_text.get("1.0", tk.END).strip(),
            }

            chapters, images = parse_docx(docx_path.read_bytes(), lang=doc_lang)
            if not chapters:
                raise ValueError(self.t("err_no_chapters"))

            cover_bytes = cover_path.read_bytes()
            cover_ext   = cover_path.suffix.lower()
            epub_bytes  = build_epub(meta, chapters, images, cover_bytes, cover_ext, lang=doc_lang)
            saved_path.write_bytes(epub_bytes)
            if self._java_available:
                epubcheck_ran, valid, messages = validate_epub(epub_bytes)
            else:
                epubcheck_ran, valid, messages = False, True, []

            result = ConversionResult(
                meta, chapters, images, epub_bytes, cover_bytes, cover_ext,
                saved_path, valid, messages, epubcheck_ran=epubcheck_ran,
            )
            preview_html_path = write_preview_html(result, self.ui_lang_var.get())
            self.after(0, lambda: self.show_result(result, preview_html_path))
        except Exception as exc:
            self.after(0, lambda: self.show_error(exc))

    def show_result(self, result: ConversionResult, preview_html_path: Path):
        self.result = result
        self.preview_html_path = preview_html_path
        self._save_profile()
        self.progress.stop()
        self._set_convert_buttons_state(tk.NORMAL)
        self.open_file_button.configure(state=tk.NORMAL)
        self.open_folder_button.configure(state=tk.NORMAL)
        self.open_html_button.configure(state=tk.NORMAL)

        if not result.epubcheck_ran:
            validation = self.t("epubcheck_skipped")
        elif result.valid:
            validation = "EpubCheck: OK"
        else:
            validation = f"EpubCheck: {len(result.messages)} " + self.t("epubcheck_messages")
        self.status_var.set(self.t("status_done") + str(result.saved_path))
        self.result_label.configure(
            text=validation + "\n" + self.t("label_file") + " " + str(result.saved_path)
        )

        self.chapter_list.delete(0, tk.END)
        for index, (_cid, title, _epub_type, _role, _body) in enumerate(result.chapters, 1):
            self.chapter_list.insert(tk.END, f"{index}. {html.unescape(title)}")
        self.chapter_list.selection_set(0)
        self.update_preview(0)

        if result.epubcheck_ran and not result.valid:
            details = "\n".join(str(msg) for msg in result.messages[:10])
            messagebox.showwarning(APP_TITLE, self.t("warn_epubcheck") + details)
        else:
            messagebox.showinfo(APP_TITLE, self.t("info_saved") + str(result.saved_path))

    def show_error(self, exc: Exception):
        self.progress.stop()
        self._set_convert_buttons_state(tk.NORMAL)
        self.status_var.set(self.t("status_error"))
        self.result_label.configure(text=self.t("label_error") + " " + str(exc))
        messagebox.showerror(APP_TITLE, str(exc))

    def on_chapter_select(self, _event):
        selection = self.chapter_list.curselection()
        if selection:
            self.update_preview(selection[0])

    def update_preview(self, index: int):
        if not self.result:
            return
        _cid, title, _epub_type, _role, body = self.result.chapters[index]
        content = html.unescape(title) + "\n" + ("=" * len(html.unescape(title))) + "\n\n" + html_to_plain_text(body)
        self.preview_text.configure(state=tk.NORMAL)
        self.preview_text.delete("1.0", tk.END)
        self.preview_text.insert(tk.END, content)
        self.preview_text.configure(state=tk.DISABLED)

    def open_epub(self):
        if self.result:
            open_path(self.result.saved_path)

    def open_output_folder(self):
        if self.result:
            open_path(self.result.saved_path.parent)

    def open_html_preview(self):
        if self.preview_html_path and self.preview_html_path.exists():
            webbrowser.open(self.preview_html_path.as_uri())

    # ------------------------------------------------------------------
    # Batch conversion
    # ------------------------------------------------------------------

    def open_batch_dialog(self):
        if not self.author_var.get().strip():
            messagebox.showwarning(APP_TITLE, self.t("batch_need_author"))
            return
        if not Path(self.cover_path.get()).is_file():
            messagebox.showwarning(APP_TITLE, self.t("batch_need_cover"))
            return
        paths = filedialog.askopenfilenames(
            title=self.t("batch_select"),
            initialdir=self._last_docx_dir or str(default_output_dir()),
            filetypes=[("DOCX", "*.docx")],
        )
        if paths:
            self._run_batch(list(paths))

    def _run_batch(self, docx_paths: list[str]):
        win = tk.Toplevel(self)
        win.title(self.t("batch_title"))
        win.configure(background=NAVY_BG)
        win.geometry("680x500")
        win.transient(self)

        frame = ttk.Frame(win, padding=16)
        frame.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frame, text=self.t("batch_title"), style="Section.TLabel").pack(anchor=tk.W, pady=(0, 8))

        list_frame = ttk.Frame(frame)
        list_frame.pack(fill=tk.BOTH, expand=True)
        status_list = tk.Listbox(
            list_frame, width=72, height=16,
            bg=NAVY_INPUT, fg=NAVY_BG, font=self._ui_font_small,
            selectbackground=NAVY_ACCENT, selectforeground="#ffffff",
            relief=tk.SOLID, borderwidth=1,
        )
        status_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=status_list.yview)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        status_list.configure(yscrollcommand=sb.set)
        for p in docx_paths:
            status_list.insert(tk.END, f"[ ] {Path(p).name}")

        progress = ttk.Progressbar(frame, mode="determinate", maximum=len(docx_paths))
        progress.pack(fill=tk.X, pady=(8, 4))
        summary_var = tk.StringVar(value="")
        ttk.Label(frame, textvariable=summary_var, style="Card.TLabel").pack(anchor=tk.W)
        close_btn = ttk.Button(frame, text=self.t("btn_ok"), command=win.destroy, state=tk.DISABLED)
        close_btn.pack(anchor=tk.E, pady=(8, 0))

        def upd(idx, symbol, text):
            status_list.delete(idx)
            status_list.insert(idx, f"{symbol} {text}")

        def worker():
            ok = err = 0
            cover_path = Path(self.cover_path.get())
            cover_ext = cover_path.suffix.lower()
            try:
                cover_bytes = cover_path.read_bytes()
            except Exception as exc:
                win.after(0, lambda: summary_var.set(f"Cover error: {exc}"))
                win.after(0, lambda: close_btn.configure(state=tk.NORMAL))
                return
            doc_lang = self.doc_lang_var.get()

            for i, docx_str in enumerate(docx_paths):
                p = Path(docx_str)
                win.after(0, lambda idx=i, n=p.name: upd(idx, "[>]", n))
                try:
                    docx_bytes = p.read_bytes()
                    issues = inspect_docx_for_epub(docx_bytes, TRANSLATIONS[self.ui_lang_var.get()])
                    if issues:
                        raise ValueError(issues[0])
                    chapters, images = parse_docx(docx_bytes, lang=doc_lang)
                    title = html.unescape(chapters[0][1]) if chapters else p.stem
                    meta = {
                        "title":       title,
                        "author":      self.author_var.get().strip(),
                        "publisher":   self.publisher_var.get().strip() or self.author_var.get().strip(),
                        "year":        self.year_var.get().strip() or str(date.today().year),
                        "isbn":        "",
                        "description": self.desc_text.get("1.0", tk.END).strip(),
                    }
                    epub_bytes = build_epub(meta, chapters, images, cover_bytes, cover_ext, lang=doc_lang)
                    out_path = p.with_suffix(".epub")
                    out_path.write_bytes(epub_bytes)
                    label = "[OK]"
                    if self._java_available:
                        ran, valid, _msgs = validate_epub(epub_bytes)
                        if ran and not valid:
                            label = "[OK*]"
                    ok += 1
                    suffix = self.t("batch_epubcheck_warn") if label == "[OK*]" else ""
                    win.after(0, lambda idx=i, n=p.name, lb=label, sf=suffix: upd(idx, lb, n + sf))
                except Exception as exc:
                    err += 1
                    msg = f"{p.name} — {exc}"
                    win.after(0, lambda idx=i, m=msg: upd(idx, "[!!]", m))
                win.after(0, lambda v=i + 1: progress.configure(value=v))

            summary = self.t("batch_summary_ok") + str(ok)
            if err:
                summary += self.t("batch_summary_err") + str(err)
            win.after(0, lambda: summary_var.set(summary))
            win.after(0, lambda: close_btn.configure(state=tk.NORMAL))

        threading.Thread(target=worker, daemon=True).start()


def main():
    app = EpubConverterApp()
    app.mainloop()


if __name__ == "__main__":
    main()
