#!/usr/bin/env python3
import csv, json, os, re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime

import ui.colors
# --- optional format support (install as needed) ---
# pip install pypdf            (PDF)
# pip install striprtf         (RTF)
# pip install python-docx      (DOCX)
# pip install ebooklib         (EPUB)
# pip install html2text        (HTML)
# LIT: convert externally with `lit2html`/Calibre; no pure-python lib.
try:
    from pypdf import PdfReader
except ImportError:
    PdfReader = None
try:
    from striprtf.striprtf import rtf_to_text
except ImportError:
    rtf_to_text = None
try:
    import docx  # python-docx
except ImportError:
    docx = None
try:
    from ebooklib import epub
    import ebooklib
except ImportError:
    epub = None
    ebooklib = None
try:
    import html2text
except ImportError:
    html2text = None



WORD_RE = re.compile(r"[A-Za-z0-9\u0400-\u04FF]+(?:-[A-Za-z0-9\u0400-\u04FF]+)*")
DICT_FILE = "dict.json"
STATE_FILE = "state.json"

import sys

def set_titlebar_color(window, hex_color: str, dark: bool = True):
    """
    Color the native Windows title bar.
    Returns True on success, False if unsupported (non-Windows / old build).
    """
    if sys.platform != "win32":
        return False
    try:
        import ctypes
        from ctypes import wintypes
    except ImportError:
        return False

    # --- helpers ---
    def _to_colorref(hex_str: str) -> int:
        # DWM wants 0x00BBGGRR, not RGB
        h = hex_str.lstrip("#")
        r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
        return (b << 16) | (g << 8) | r

    hwnd = wintypes.HWND(window.winfo_id())
    # we need the *top-level* window, not the child window Tk gives us
    user32 = ctypes.windll.user32
    hwnd = wintypes.HWND(user32.GetAncestor(hwnd, 2))  # GA_ROOT = 2

    dwm = ctypes.windll.dwmapi
    colorref = ctypes.c_int(_to_colorref(hex_color))

    # DWMWA_USE_IMMERSIVE_DARK_MODE = 20 (Win10 20H1+ / Win11)
    # older builds used 19; try both
    dark_flag = ctypes.c_int(1 if dark else 0)
    for attr in (20, 19):
        res = dwm.DwmSetWindowAttribute(
            hwnd, attr, ctypes.byref(dark_flag), ctypes.sizeof(dark_flag))
        if res == 0:
            break

    # DWMWA_CAPTION_COLOR = 35 (Win11 22H2+ only)
    # DWMWA_BORDER_COLOR  = 34
    try:
        dwm.DwmSetWindowAttribute(
            hwnd, 35, ctypes.byref(colorref), ctypes.sizeof(colorref))
        dwm.DwmSetWindowAttribute(
            hwnd, 34, ctypes.byref(colorref), ctypes.sizeof(colorref))
    except Exception:
        pass

    # force a redraw so the change takes effect immediately
    user32.SetWindowPos(hwnd, 0, 0, 0, 0, 0, 0x0027)  # SWP_NOMOVE|SIZE|ZORDER|NOACTIVATE
    return True

def extract_words(text: str):
    return [m.group(0).lower() for m in WORD_RE.finditer(text)]

def _read_text_bytes(path: str) -> str:
    """Read a file trying several encodings (UTF-8, UTF-16, cp1251, latin-1)."""
    with open(path, "rb") as f:
        raw = f.read()
    for enc in ("utf-8-sig", "utf-8", "utf-16", "cp1251", "latin-1"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


def extract_text_from_file(path: str) -> str:
    """
    Extract plain text from a file based on its extension.
    Raises ValueError with a helpful message when a format cannot be read.
    """
    ext = os.path.splitext(path)[1].lower()

    # --- plain text family ---
    if ext in ("", ".txt", ".text", ".log", ".md", ".markdown",
               ".csv", ".tsv", ".json", ".xml", ".yml", ".yaml",
               ".py", ".js", ".ts", ".c", ".h", ".cpp", ".java",
               ".cs", ".go", ".rs", ".rb", ".php", ".sh", ".bat",
               ".ini", ".cfg", ".conf", ".srt", ".tex"):
        return _read_text_bytes(path)

    # --- RTF ---
    if ext == ".rtf":
        if rtf_to_text is None:
            raise ValueError(
                "RTF support requires the 'striprtf' package.\n"
                "Install it with:  pip install striprtf")
        raw = _read_text_bytes(path)
        return rtf_to_text(raw)

    # --- PDF ---
    if ext == ".pdf":
        if PdfReader is None:
            raise ValueError(
                "PDF support requires the 'pypdf' package.\n"
                "Install it with:  pip install pypdf")
        reader = PdfReader(path)
        parts = []
        for page in reader.pages:
            try:
                parts.append(page.extract_text() or "")
            except Exception:
                parts.append("")
        return "\n".join(parts)

    # --- DOCX ---
    if ext == ".docx":
        if docx is None:
            raise ValueError(
                "DOCX support requires the 'python-docx' package.\n"
                "Install it with:  pip install python-docx")
        d = docx.Document(path)
        parts = [p.text for p in d.paragraphs]
        for table in d.tables:
            for row in table.rows:
                for cell in row.cells:
                    parts.append(cell.text)
        return "\n".join(parts)

    # --- EPUB ---
    if ext == ".epub":
        if epub is None or ebooklib is None:
            raise ValueError(
                "EPUB support requires the 'EbookLib' package.\n"
                "Install it with:  pip install EbookLib")
        book = epub.read_epub(path)
        parts = []
        for item in book.get_items_of_type(ebooklib.ITEM_DOCUMENT):
            try:
                html = item.get_content().decode("utf-8", errors="replace")
            except Exception:
                continue
            if html2text is not None:
                h = html2text.HTML2Text()
                h.ignore_links = True
                h.ignore_images = True
                parts.append(h.handle(html))
            else:
                # crude fallback: strip tags
                parts.append(re.sub(r"<[^>]+>", " ", html))
        return "\n".join(parts)

    # --- HTML ---
    if ext in (".html", ".htm", ".xhtml"):
        raw = _read_text_bytes(path)
        if html2text is not None:
            h = html2text.HTML2Text()
            h.ignore_links = True
            h.ignore_images = True
            return h.handle(raw)
        return re.sub(r"<[^>]+>", " ", raw)

    # --- LIT (Microsoft Reader) ---
    # No pure-python reader. Convert with Calibre first, e.g.:
    #   ebook-convert book.lit book.txt
    if ext == ".lit":
        raise ValueError(
            "LIT files are not supported directly.\n"
            "Convert the .lit file to .txt or .html first, for example with "
            "Calibre:\n\n    ebook-convert \"book.lit\" \"book.txt\"\n\n"
            "Then open the resulting file here.")

    # --- unknown: try reading as text anyway ---
    try:
        return _read_text_bytes(path)
    except OSError as e:
        raise ValueError(f"Unsupported file type '{ext}': {e}")

def load_translation_dict():

    def resource_path(relative_path):
        """ Get the absolute path to the resource, compatible with development and PyInstaller packaging modes """
        if hasattr(sys, '_MEIPASS'):
            # Running in PyInstaller packaged environment
            return os.path.join(sys._MEIPASS, relative_path)
        # Running in normal development environment
        return os.path.join(os.path.abspath("."), relative_path)


    try:
        with open(resource_path(DICT_FILE), "r", encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, dict):
            raise ValueError("dict.json must contain a JSON object")
        return {str(k).lower(): str(v) for k, v in data.items()}
    except (OSError, ValueError, json.JSONDecodeError) as e:
        messagebox.showwarning(
            "dict.json",
            f'Failed to load "{DICT_FILE}":\n{e}\n\nTranslations will be empty.')
        return {}


def load_state():
    """Load state.json. Missing or broken file -> empty dict."""
    if not os.path.exists(STATE_FILE):
        return {}
    try:
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, dict):
            raise ValueError("state.json must contain a JSON object")
        return data
    except (OSError, ValueError, json.JSONDecodeError):
        return {}


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Text—>Dictionary")

        # --- data ---
        self.dictionary = {}          # word -> count
        self.learned = set()
        self.translations = load_translation_dict()   # read-only lookup
        self.selected = set()
        self.last_click_index = -1
        self.last_click_word = None
        self.last_occurrence = 0
        self.current_tab = "all"
        self.rows_data = []
        self._drag_active = False

        # --- zoom ---
        self.scale = 1.0
        self._base_geom = (1400.0, 720.0)   # unscaled window size
        self._cols_base = ()                # filled in _build_ui

        # theme: default light, restored from state.json later in _restore_state
        self.theme_name = "light"
        self.palette = ui.colors.LIGHT

        self._setup_style()
        self._build_ui()
        self._restore_state()
        self._refresh()

        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ------------------------------------------------------------ STYLE
    def _px(self, n):
        """Scale a pixel value by the current zoom factor."""
        return max(1, int(round(n * self.scale)))
        
    def _setup_style(self):
        s = ttk.Style(self)
        try:
            s.theme_use("clam")
        except tk.TclError:
            pass

        p = self.palette
        self.configure(bg=p["bg"])

        s.configure(".", background=p["bg"], foreground=p["text"],
                    font=("Segoe UI", self._px(10)))
        s.configure("TFrame", background=p["bg"])
        s.configure("Card.TFrame", background=p["card"])

        s.configure("TLabel", background=p["bg"], foreground=p["text"])
        s.configure("Stats.TLabel", foreground=p["muted"],
                    font=("Segoe UI", self._px(9)))
        s.configure("Title.TLabel",
                    font=("Segoe UI Semibold", self._px(11)),
                    foreground=p["text"])

        # Buttons
        s.configure("TButton", padding=(self._px(10), self._px(6)),
                    relief="flat", background=p["card"], foreground=p["text"],
                    borderwidth=1)
        s.map("TButton",
            background=[("active", p["accent_soft"]),
                        ("pressed", p["accent_soft2"])],
            foreground=[("disabled", p["muted"])],
            bordercolor=[("!disabled", p["border"])])

        s.configure("Primary.TButton", background=p["accent"], foreground="white",
                    font=("Segoe UI Semibold", self._px(10)))
        s.map("Primary.TButton",
            background=[("active", p["accent2"]), ("pressed", p["accent3"])])

        s.configure("Danger.TButton", background=p["danger_bg"],
                    foreground=p["danger"])
        s.map("Danger.TButton", background=[("active", p["danger_bg2"])])

        # Tabs (segmented)
        s.configure("Tab.TButton", padding=(self._px(14), self._px(7)),
                    relief="flat", background=p["tab_bg"], foreground=p["muted"],
                    font=("Segoe UI", self._px(10)))
        s.configure("TabActive.TButton", padding=(self._px(14), self._px(7)),
                    relief="flat", background=p["card"], foreground=p["accent"],
                    font=("Segoe UI Semibold", self._px(10)))

        # Treeview
        s.configure("Treeview", background=p["card"], fieldbackground=p["card"],
                    foreground=p["text"], rowheight=self._px(28), borderwidth=0,
                    font=("Segoe UI", self._px(10)))
        s.configure("Treeview.Heading", background=p["head_bg"],
                    foreground=p["text"], relief="flat",
                    font=("Segoe UI Semibold", self._px(10)),
                    padding=(self._px(8), self._px(6)))
        s.map("Treeview",
            background=[("selected", p["accent_soft2"])],
            foreground=[("selected", p["text"])])

        # Scrollbars
        s.configure("Vertical.TScrollbar", background=p["scroll"],
                    troughcolor=p["bg"], borderwidth=0,
                    arrowsize=self._px(12))
        s.map("Vertical.TScrollbar",
            background=[("active", p["muted"])])

        # Paned window
        s.configure("TPanedwindow", background=p["bg"])
        
    def toggle_theme(self):
        self.set_theme("dark" if self.theme_name == "light" else "light")
    
    def set_theme(self, name):
        if name not in ("light", "dark"):
            name = "light"

        self.theme_name = name
        self.palette = ui.colors.DARK if name == "dark" else ui.colors.LIGHT

        # remember current geometry and text content so the rebuild is lossless
        geom = self.geometry()
        old_text = self.text.get("1.0", "end-1c")   # <-- save text here
        yview = self.text.yview()                   # <-- optional: keep scroll pos

        # tear down all children of the root window
        for child in list(self.winfo_children()):
            child.destroy()

        self._setup_style()
        self._build_ui()
        self.geometry(geom)

        # restore text and scroll position
        if old_text:
            self.text.insert("1.0", old_text)
            try:
                self.text.yview_moveto(yview[0])
            except tk.TclError:
                pass

        # redraw table from existing data
        self._refresh()
        self._update_theme_button()
        self.update_idletasks()
        set_titlebar_color(self, self.palette["bg"], dark=(name == "dark"))

    def _update_theme_button(self):
        if not hasattr(self, "theme_btn"):
            return
        if self.theme_name == "dark":
            self.theme_btn.config(text="☀ Light")
        else:
            self.theme_btn.config(text="🌙 Dark")
    # ------------------------------------------------------------------ UI
    def _build_ui(self):
        # ---- top bar ----
        bar = ttk.Frame(self, padding=(self._px(12), self._px(10)))
        bar.pack(side=tk.TOP, fill=tk.X)
        ttk.Label(bar, text="📚  Text Scanner", style="Title.TLabel").pack(side=tk.LEFT)
        
        # store the reference so we can update the icon when theme changes
        self.theme_btn = ttk.Button(bar, text="🌙 Dark",
                                    command=self.toggle_theme)
        self.theme_btn.pack(side=tk.RIGHT, padx=(0, self._px(6)))
        
        ttk.Button(bar, text="Choose File", style="Primary.TButton",
                command=self.load_text_file).pack(side=tk.RIGHT)
        # separator line
        tk.Frame(self, bg="#d8dee7", height=1).pack(fill=tk.X)
        self.paned = tk.PanedWindow(self, orient=tk.HORIZONTAL,
                                    bg=self.palette["bg"],
                                    sashwidth=self._px(6),
                                    sashrelief=tk.FLAT, bd=0)
        
        self.paned.pack(fill=tk.BOTH, expand=True,
                        padx=self._px(12), pady=self._px(12))
        # ================= LEFT: text card =================
        left = ttk.Frame(self.paned, style="Card.TFrame", padding=1)
        self.paned.add(left, minsize=self._px(300), stretch="always")
        text_frame = ttk.Frame(left, style="Card.TFrame")
        text_frame.pack(fill=tk.BOTH, expand=True)
        self.text = tk.Text(text_frame, wrap=tk.WORD, undo=True,
                relief="flat", bd=0,
                padx=self._px(10), pady=self._px(8),
                font=("Consolas", self._px(10)),
                bg=self.palette["text_widget_bg"],
                fg=self.palette["text_widget_fg"],
                insertbackground=self.palette["insert"],
                selectbackground=self.palette["accent_soft2"])
        text_vsb = ttk.Scrollbar(text_frame, orient="vertical",
                                command=self.text.yview)
        self.text.configure(yscrollcommand=text_vsb.set)
        self.text.grid(row=0, column=0, sticky="nsew")
        text_vsb.grid(row=0, column=1, sticky="ns")
        text_frame.rowconfigure(0, weight=1)
        text_frame.columnconfigure(0, weight=1)
        self.text.tag_configure("highlight", background=self.palette["highlight"])
        self.text.bind("<MouseWheel>", self._on_text_wheel)
        self.text.bind("<Button-4>", self._on_text_wheel)
        self.text.bind("<Button-5>", self._on_text_wheel)
        # ================= RIGHT: table card =================
        right = ttk.Frame(self.paned, style="Card.TFrame")
        self.paned.add(right, minsize=self._px(400), stretch="always")
        # stats line
        stats_bar = ttk.Frame(right, style="Card.TFrame",
                            padding=(self._px(10), self._px(6),
                                    self._px(10), 0))
        stats_bar.pack(fill=tk.X)
        self.stats_var = tk.StringVar(value="Dictionary is empty — choose a file to scan it")
        ttk.Label(stats_bar, textvariable=self.stats_var,
                style="Stats.TLabel").pack(side=tk.LEFT)
        # tab bar + load/save on one line
        tabbar = ttk.Frame(right, style="Card.TFrame",
                        padding=(self._px(10), self._px(8),
                                    self._px(10), 0))
        tabbar.pack(fill=tk.X)
        # right-side buttons first (so they stay right when packing left items)
        ttk.Button(tabbar, text="Save table",
                command=self.save_csv).pack(side=tk.RIGHT)
        ttk.Button(tabbar, text="Load table",
                command=self.load_csv).pack(side=tk.RIGHT,
                                                padx=(0, self._px(4)))
        self.tab_buttons = {}
        for key, label in (("all", "All words"),
                        ("learned", "Learned")):
            b = ttk.Button(tabbar, text=label, style="Tab.TButton",
                        command=lambda k=key: self.switch_tab(k))
            b.pack(side=tk.LEFT, padx=(0, self._px(4)))
            self.tab_buttons[key] = b
        self.selbar = tk.Frame(right, bg=self.palette["sel_bg"],
                            bd=1, relief=tk.SOLID,
                            highlightbackground=self.palette["border"])
        self.sel_label = tk.Label(self.selbar, text="", bg=self.palette["sel_bg"],
                                fg=self.palette["text"], anchor="w",
                                font=("Segoe UI", self._px(9)))
        self.sel_label.pack(side=tk.LEFT, padx=self._px(8))
        self.sel_actions = tk.Frame(self.selbar, bg=self.palette["sel_bg"])
        self.sel_actions.pack(side=tk.LEFT)
        # table
        table_frame = ttk.Frame(right, style="Card.TFrame",
                                padding=(self._px(10), self._px(6),
                                        self._px(10), self._px(10)))
        table_frame.pack(fill=tk.BOTH, expand=True)
        cols = ("word", "count", "translate", "actions")
        self.tree = ttk.Treeview(table_frame, columns=cols, show="headings",
                                selectmode="extended")
        self._cols_base = (
            ("word",      160, "Word",         "w"),
            ("count",      90, "Occurrences",  "center"),
            ("translate", 260, "Translate",    "w"),
            ("actions",   210, "Actions",      "w"),
        )
        for c, w, label, anchor in self._cols_base:
            self.tree.heading(c, text=label)
            self.tree.column(c, width=self._px(w), anchor=anchor)
        vsb = ttk.Scrollbar(table_frame, orient="vertical",
                            command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        self.tree.bind("<MouseWheel>", self._on_tree_wheel)
        self.tree.bind("<Button-4>", self._on_tree_wheel)
        self.tree.bind("<Button-5>", self._on_tree_wheel)
        self.tree.bind("<ButtonPress-1>", self._on_mouse_down)
        self.tree.bind("<B1-Motion>", self._on_mouse_drag)
        self.tree.bind("<ButtonRelease-1>", self._on_mouse_up)
        self.tree.bind("<Motion>", self._on_motion)
        self.tree.tag_configure("untranslated", foreground=self.palette["untrans"])
        self.tree.tag_configure("odd", background=self.palette["zebra"])
        self.tree.tag_configure("even", background=self.palette["row"])
        self._bind_zoom_keys()
    # ------------------------------------------------------------ zoom
    def _bind_zoom_keys(self):
        """Ctrl + / Ctrl - / Ctrl 0 work no matter which widget has focus."""
        for seq in ("<Control-plus>", "<Control-equal>", "<Control-KP_Add>"):
            self.bind_all(seq, lambda e: self._zoom_by(+1), add="+")
        for seq in ("<Control-minus>", "<Control-KP_Subtract>"):
            self.bind_all(seq, lambda e: self._zoom_by(-1), add="+")
        for seq in ("<Control-0>", "<Control-KP_0>"):
            self.bind_all(seq, lambda e: self._zoom_reset(), add="+")
    def _zoom_by(self, direction):
        self._set_scale(round(self.scale + direction * 0.1, 2))
        return "break"
    def _zoom_reset(self):
        self._set_scale(1.0)
        return "break"
    def _set_scale(self, value):
        value = max(0.5, min(3.0, value))
        if abs(value - self.scale) < 1e-6:
            return
        # Remember the "unscaled" size from the current real size.
        try:
            w, h = self.winfo_width(), self.winfo_height()
            if w > 1 and h > 1 and self.scale > 0:
                self._base_geom = (w / self.scale, h / self.scale)
        except tk.TclError:
            pass
        self.scale = value
        self._apply_scale()
    def _apply_scale(self):
        """Re-apply scaled font sizes and pixel dimensions to every widget."""
        self._setup_style()
        self.text.configure(
            font=("Consolas", self._px(10)),
            padx=self._px(10), pady=self._px(8))
        for c, w, _label, _anchor in self._cols_base:
            self.tree.column(c, width=self._px(w))
        self.paned.configure(sashwidth=self._px(6))
        self.paned.pack_configure(padx=self._px(12), pady=self._px(12))
        panes = self.paned.panes()
        if len(panes) >= 2:
            self.paned.paneconfigure(panes[0], minsize=self._px(300))
            self.paned.paneconfigure(panes[1], minsize=self._px(400))
        self.sel_label.configure(font=("Segoe UI", self._px(9)))
    # --------------------------------------------------- mousewheel
    def _on_text_wheel(self, event):
        if event.num == 4 or event.delta > 0:
            self.text.yview_scroll(-1, "units")
        elif event.num == 5 or event.delta < 0:
            self.text.yview_scroll(1, "units")
        return "break"
    def _on_tree_wheel(self, event):
        if event.num == 4 or event.delta > 0:
            self.tree.yview_scroll(-1, "units")
        elif event.num == 5 or event.delta < 0:
            self.tree.yview_scroll(1, "units")
        return "break"
    # -------------------------------------------------------- tab logic
    def switch_tab(self, tab):
        self.current_tab = tab
        self.selected.clear()
        self.last_click_index = -1
        self.last_click_word = None
        self.last_occurrence = 0
        self._refresh()
    def _update_tab_labels(self):
        n_learned = len(self.learned)
        labels = {
            "all":     "All words",
            "learned": f"Learned ({n_learned})" if n_learned else "Learned",
        }
        for key, b in self.tab_buttons.items():
            b.config(text=labels[key],
                    style="TabActive.TButton" if key == self.current_tab
                    else "Tab.TButton")
    # -------------------------------------------------------- rendering
    def _filtered_entries(self):
        if self.current_tab == "all":
            pool = ((w, c) for w, c in self.dictionary.items()
                    if w not in self.learned)
        else:
            pool = ((w, c) for w, c in self.dictionary.items() if w in self.learned)
        return sorted(pool, key=lambda x: (-x[1], x[0]))
    def _refresh(self):
        self._update_tab_labels()
        self.tree.delete(*self.tree.get_children())
        entries = self._filtered_entries()
        self.rows_data = entries
        self.tree["displaycolumns"] = ("word", "count", "translate", "actions")
        for i, (w, c) in enumerate(entries):
            zebra = "odd" if i % 2 else "even"
            tr = self.translations.get(w, "")
            values = (w, c, tr if tr else "—", self._action_label_for(w))
            tags = (zebra,) if tr else (zebra, "untranslated")
            iid = self.tree.insert("", "end", iid=f"row{i}",
                                values=values, tags=tags)
            if w in self.selected:
                self.tree.selection_add(iid)
        self._update_stats()
        self._update_selection_ui()
    def _update_row(self, word):
        iid = None
        for child in self.tree.get_children():
            if self.tree.item(child, "values")[0] == word:
                iid = child
                break
        if iid is None:
            return
        c = self.dictionary.get(word, 0)
        # preserve zebra striping from current position
        idx = int(iid.replace("row", ""))
        zebra = "odd" if idx % 2 else "even"
        tr = self.translations.get(word, "")
        self.tree.item(
            iid,
            values=(word, c, tr if tr else "—", self._action_label_for(word)),
            tags=(zebra,) if tr else (zebra, "untranslated"),
        )
        self._update_stats()
        self._update_selection_ui()
    def _remove_row(self, word):
        for child in self.tree.get_children():
            if self.tree.item(child, "values")[0] == word:
                self.tree.delete(child)
                break
        self.rows_data = [r for r in self.rows_data if r[0] != word]
        self._update_stats()
        self._update_selection_ui()
    def _action_label_for(self, w):
        # Buttons are rendered as bracketed labels inside the Actions cell.
        if self.current_tab == "all":
            return "[✓ I know]"
        return "[↩ Return]"
    def _action_buttons_for(self, w):
        if self.current_tab == "all":
            return [("✓ I know", self.learn_word)]
        return [("↩ Return", self.unlearn_word)]
    def _update_stats(self):
        if not self.dictionary:
            self.stats_var.set("Dictionary is empty — choose a file to scan it")
            return
        total_unique = len(self.dictionary)
        total_occur = sum(self.dictionary.values())
        visible = self._filtered_entries()
        vis_unique = len(visible)
        vis_occur = sum(c for _, c in visible)
        learned_occur = sum(self.dictionary[w] for w in self.learned)
        untranslated = sum(1 for w, _ in visible if w not in self.translations)
        parts = [f"Showing: {vis_unique} of {total_unique} words",
                f"occurrences: {vis_occur} of {total_occur}"]
        if self.learned:
            parts.append(f"known: {len(self.learned)} words ({learned_occur} occurrences)")
        if untranslated:
            parts.append(f"untranslated: {untranslated}")
        self.stats_var.set(" · ".join(parts))
    # -------------------------------------------------- selection toolbar
    def _update_selection_ui(self):
        if len(self.selected) <= 1:
            self.selbar.pack_forget()
            return
        self.selbar.pack(fill=tk.X, pady=(self._px(4), 0),
                        before=self.tree.master)
        self.sel_label.config(text=f"{len(self.selected)} selected")
        for child in self.sel_actions.winfo_children():
            child.destroy()
        if self.current_tab == "all":
            ttk.Button(self.sel_actions, text="✓ I know all",
                    command=self.bulk_learn).pack(side=tk.LEFT,
                                                    padx=self._px(2))
        else:
            ttk.Button(self.sel_actions, text="↩ Return all",
                    command=self.bulk_unlearn).pack(side=tk.LEFT,
                                                        padx=self._px(2))
    def clear_selection(self):
        self.selected.clear()
        for iid in self.tree.selection():
            self.tree.selection_remove(iid)
        self._update_selection_ui()
    # ---------------------------------------------------- mouse handling
    def _row_word(self, iid):
        return self.tree.item(iid, "values")[0]
    def _actions_column_index(self):
        return "#4"
    def _hit_action_button(self, event, iid):
        """Return (label, callback) if the click landed on an action button, else None."""
        if self.tree.identify_column(event.x) != self._actions_column_index():
            return None
        # cell bounding box
        bbox = self.tree.bbox(iid, column=self._actions_column_index())
        if not bbox:
            return None
        cx, cy, cw, ch = bbox
        rel_x = event.x - cx
        if rel_x < 0 or rel_x > cw:
            return None
        word = self._row_word(iid)
        buttons = self._action_buttons_for(word)
        # measure rendered widths using the same font as the tree
        import tkinter.font as tkfont
        f = tkfont.Font(font=("Segoe UI", self._px(10)))
        pad = self._px(8)
        x_cursor = 2
        for label, cb in buttons:
            text_w = f.measure(f" {label} ") + pad
            if x_cursor <= rel_x <= x_cursor + text_w:
                return label, cb, word
            x_cursor += text_w + self._px(4)
        return None
    def _on_motion(self, event):
        """Show hand cursor over action buttons."""
        iid = self.tree.identify_row(event.y)
        if iid and self._hit_action_button(event, iid):
            self.tree.configure(cursor="hand2")
        else:
            self.tree.configure(cursor="")
    def _on_mouse_down(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell":
            self._drag_active = False
            return
        iid = self.tree.identify_row(event.y)
        if not iid:
            self._drag_active = False
            return
        hit = self._hit_action_button(event, iid)
        if hit:
            self._drag_active = False
            _label, callback, word = hit
            callback(word)
            return
        if self.tree.identify_column(event.x) == self._actions_column_index():
            # clicked action cell but not on a button — do nothing
            self._drag_active = False
            return
        self._drag_active = True
        index = int(iid.replace("row", ""))
        if event.state & 0x0004:  # Ctrl
            word = self._row_word(iid)
            if word in self.selected:
                self.selected.discard(word)
                self.tree.selection_remove(iid)
            else:
                self.selected.add(word)
                self.tree.selection_add(iid)
            self.last_click_index = index
            self._update_selection_ui()
            return
        if event.state & 0x0001:  # Shift
            if self.last_click_index >= 0:
                lo, hi = sorted((self.last_click_index, index))
                for i in range(lo, hi + 1):
                    iid2 = f"row{i}"
                    if self.tree.exists(iid2):
                        self.selected.add(self._row_word(iid2))
                        self.tree.selection_add(iid2)
                self._update_selection_ui()
            return
        # plain click: clear selection, select this row
        self.selected.clear()
        for s in self.tree.selection():
            self.tree.selection_remove(s)
        word = self._row_word(iid)
        self.selected.add(word)
        self.tree.selection_add(iid)
        if index == self.last_click_index and word == self.last_click_word:
            # same row clicked again → advance to next occurrence in the text
            positions = self._find_word_positions(word)
            if positions:
                self.last_occurrence = (self.last_occurrence + 1) % len(positions)
                self._jump_to_word(word, self.last_occurrence)
        else:
            # new row → focus first occurrence
            self.last_occurrence = 0
            self._jump_to_word(word, 0)
        self.last_click_index = index
        self.last_click_word = word
        self._update_selection_ui()
    def _on_mouse_drag(self, event):
        if not self._drag_active:
            return
        iid = self.tree.identify_row(event.y)
        if not iid:
            return
        index = int(iid.replace("row", ""))
        if self.last_click_index < 0:
            self.last_click_index = index
        lo, hi = sorted((self.last_click_index, index))
        for i in range(lo, hi + 1):
            iid2 = f"row{i}"
            if self.tree.exists(iid2):
                self.selected.add(self._row_word(iid2))
                self.tree.selection_add(iid2)
        self._update_selection_ui()
    def _on_mouse_up(self, event):
        self._drag_active = False
    # -------------------------------------------------------- actions
    def learn_word(self, w):
        self.learned.add(w)
        self._remove_row(w)
    def unlearn_word(self, w):
        self.learned.discard(w)
        self._remove_row(w)
    def bulk_learn(self):
        words = list(self.selected)
        for w in words:
            self.learned.add(w)
        self.selected.clear()
        self._refresh()
    def bulk_unlearn(self):
        for w in list(self.selected):
            self.learned.discard(w)
        self.selected.clear()
        self._refresh()
    # ------------------------------------------------------- text jump
    def _find_word_positions(self, word):
        """Return list of (start, end) text indices for whole-word matches of `word`."""
        positions = []
        count = tk.IntVar()
        start = "1.0"
        while True:
            pos = self.text.search(word, start, stopindex="end",
                                nocase=True, count=count)
            if not pos or count.get() == 0:
                break
            end = f"{pos}+{count.get()}c"
            before = self.text.get(f"{pos}-1c", pos) if pos != "1.0" else ""
            after = self.text.get(end, f"{end}+1c")
            if (not before or not re.match(r"[A-Za-z0-9\u0400-\u04FF-]", before)) and \
            (not after or not re.match(r"[A-Za-z0-9\u0400-\u04FF-]", after)):
                positions.append((pos, end))
            start = end
        return positions
    def _jump_to_word(self, word, occurrence=0):
        """Highlight all occurrences, scroll to and focus the given occurrence."""
        self.text.tag_remove("highlight", "1.0", "end")
        positions = self._find_word_positions(word)
        if not positions:
            return
        for pos, end in positions:
            self.text.tag_add("highlight", pos, end)
        # clamp to available occurrences
        idx = occurrence % len(positions)
        pos, end = positions[idx]
        self.text.see(pos)
        self.text.mark_set(tk.INSERT, pos)
        self.text.tag_remove(tk.SEL, "1.0", "end")
        self.text.tag_add(tk.SEL, pos, end)
        self.text.focus_set()
    # ------------------------------------------------------- IO: text
    def load_text_file(self):
        path = filedialog.askopenfilename(
            filetypes=[
                ("Books & documents",
                "*.txt *.text *.md *.markdown *.rtf *.pdf *.docx "
                "*.epub *.html *.htm *.xhtml *.lit *.log *.srt"),
                ("Text files", "*.txt *.text *.log *.md"),
                ("RTF", "*.rtf"),
                ("PDF", "*.pdf"),
                ("DOCX", "*.docx"),
                ("EPUB", "*.epub"),
                ("HTML", "*.html *.htm *.xhtml"),
                ("LIT (convert first)", "*.lit"),
                ("All files", "*.*"),
            ])
        if not path:
            return
        try:
            text = extract_text_from_file(path)
        except ValueError as e:
            messagebox.showerror("Open", str(e))
            return
        except OSError as e:
            messagebox.showerror("Error", f"Failed to read file:\n{e}")
            return
        if not text or not text.strip():
            messagebox.showwarning(
                "Open",
                f'No text could be extracted from:\n{os.path.basename(path)}\n\n'
                "If this is a scanned PDF, OCR is required first.")
            return
        self.text.delete("1.0", "end")
        self.text.insert("1.0", text)
        words = extract_words(text)
        if not words:
            messagebox.showwarning("Scan", "No words found in the text.")
            return
        for w in words:
            self.dictionary[w] = self.dictionary.get(w, 0) + 1
        self._refresh()
    # ------------------------------------------------------- IO: CSV
    def load_csv(self):
        path = filedialog.askopenfilename(
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8", newline="") as f:
                reader = csv.reader(f)
                for row in reader:
                    if not row: continue
                    if row[0].strip().lower() == "word": continue
                    w = row[0].strip().lower()
                    if not w: continue
                    try:
                        c = int(row[1]) if len(row) > 1 and row[1].strip() else 0
                    except ValueError:
                        c = 0
                    # keep the larger of existing vs CSV count so real
                    # occurrence counts from a text scan are not clobbered
                    self.dictionary[w] = max(self.dictionary.get(w, 0), c)
                    if self.current_tab == "learned":
                        self.learned.add(w)
                    else:
                        self.learned.discard(w)
        except OSError as e:
            messagebox.showerror("Error",
                f'Failed to load "{os.path.basename(path)}":\n{e}')
            return
        self._refresh()
    def save_csv(self):
        if not self.dictionary:
            messagebox.showinfo("Save", "Nothing to save — the dictionary is empty.")
            return
        if self.current_tab == "all":
            rows = [(w, c) for w, c in self.dictionary.items()
                    if w not in self.learned]
        else:
            rows = [(w, self.dictionary[w]) for w in self.learned]
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            initialfile=f"{self.current_tab}_{datetime.now().strftime('%Y-%m-%d-%H-%M-%S')}.csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8", newline="") as f:
                writer = csv.writer(f)
                writer.writerow(["word", "occurrences", "translate"])
                for w, c in sorted(rows, key=lambda x: (-x[1], x[0])):
                    writer.writerow([w, c, self.translations.get(w, "")])
        except OSError as e:
            messagebox.showerror("Error", f"Failed to save file:\n{e}")
            return
        messagebox.showinfo("Saved", f"Saved {self.current_tab} CSV to:\n{path}")
    # ------------------------------------------------------- state
    def _restore_state(self):
        state = load_state()
            # theme first — it may rebuild the UI
        saved_theme = state.get("theme", "light")
        if saved_theme in ("light", "dark") and saved_theme != self.theme_name:
            self.set_theme(saved_theme)
        else:
            self._update_theme_button()

        if not state:
            self.geometry("1400x720")
            return
        geom = state.get("geometry")
        if isinstance(geom, str) and geom:
            try:
                self.geometry(geom)
            except tk.TclError:
                self.geometry("1400x720")
        else:
            self.geometry("1400x720")
        text = state.get("text", "")
        if isinstance(text, str) and text:
            self.text.insert("1.0", text)
        raw_dict = state.get("dictionary", {})
        if isinstance(raw_dict, dict):
            for w, c in raw_dict.items():
                try:
                    self.dictionary[str(w).lower()] = int(c)
                except (TypeError, ValueError):
                    self.dictionary[str(w).lower()] = 1
        raw_learned = state.get("learned", [])
        if isinstance(raw_learned, list):
            self.learned = {str(w).lower() for w in raw_learned}
        tab = state.get("current_tab", "all")
        if tab in ("all", "learned"):
            self.current_tab = tab

    def _save_state(self):
        state = {
            "geometry": self.geometry(),
            "text": self.text.get("1.0", "end-1c"),
            "dictionary": self.dictionary,
            "learned": sorted(self.learned),
            "current_tab": self.current_tab,
            "theme": self.theme_name,     
        }
        try:
            with open(STATE_FILE, "w", encoding="utf-8") as f:
                json.dump(state, f, ensure_ascii=False, indent=2)
        except OSError as e:
            messagebox.showwarning(
                "State",
                f'Failed to save "{STATE_FILE}":\n{e}')
            
    def _on_close(self):
        self._save_state()
        self.destroy()


if __name__ == "__main__":
    App().mainloop()