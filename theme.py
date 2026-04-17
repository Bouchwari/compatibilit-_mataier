"""
theme.py — Colors, fonts, reusable widgets
"""
import tkinter as tk
from tkinter import ttk, font as tkfont

# ─── COLOR PALETTE ───────────────────────────────────────────────
BG_DARK     = "#0D1B2A"
BG_SIDEBAR  = "#0D1B2A"
BG_MAIN     = "#F0F4F8"
BG_CARD     = "#FFFFFF"
BG_HEADER   = "#1B6CA8"
BG_ROW_ALT = "#F5F7FA"
BG_INPUT    = "#FFFFFF"
BG_HOVER    = "#E8F0FE"

ACCENT      = "#1B6CA8"
ACCENT2     = "#1B998B"
ACCENT3     = "#5C3E8F"
GREEN       = "#217A3C"
GREEN_LT    = "#E2EFDA"
ORANGE      = "#C55A11"
ORANGE_LT   = "#FCE4D6"
RED         = "#C00000"
RED_LT      = "#FFE6E6"
YELLOW_LT   = "#FFEB9C"
GOLD        = "#B8860B"
GOLD_LT     = "#FFFACD"

TEXT_DARK   = "#0D1B2A"
TEXT_MID    = "#444444"
TEXT_LIGHT  = "#888888"
TEXT_WHITE  = "#FFFFFF"
TEXT_BLUE   = "#1B6CA8"

# Category colors
CAT_COLORS = {
    "مكتبيات":       "#1B6CA8",
    "ادوات التدريس": "#3D5A80",
    "ادوات النظافة": "#6B4226",
    "الصيانة":       "#2E4057",
    "مختلفات":       "#5C3E8F",
}

FONT_FAMILY = "Arial"  # supports Arabic on Windows/Linux

def F(size=11, bold=False, italic=False):
    w = "bold" if bold else "normal"
    sl = "italic" if italic else "roman"
    return (FONT_FAMILY, size, w, sl)

# ─── STYLE SETUP ─────────────────────────────────────────────────
def apply_style(root):
    style = ttk.Style(root)
    style.theme_use("clam")

    style.configure(".",
        background=BG_MAIN, foreground=TEXT_DARK,
        font=(FONT_FAMILY, 11), fieldbackground=BG_INPUT)

    style.configure("Treeview",
        background=BG_CARD, foreground=TEXT_DARK,
        rowheight=30, fieldbackground=BG_CARD,
        font=(FONT_FAMILY, 10))
    style.configure("Treeview.Heading",
        background=BG_DARK, foreground=TEXT_WHITE,
        font=(FONT_FAMILY, 10, "bold"), relief="flat")
    style.map("Treeview",
        background=[("selected", ACCENT)],
        foreground=[("selected", TEXT_WHITE)])
    style.map("Treeview.Heading",
        background=[("active", "#2E75B6")])

    style.configure("Card.TFrame",  background=BG_CARD, relief="flat")
    style.configure("Dark.TFrame",  background=BG_DARK)
    style.configure("Main.TFrame",  background=BG_MAIN)
    style.configure("Header.TFrame",background=ACCENT)

    style.configure("TCombobox",
        fieldbackground=BG_INPUT, background=BG_INPUT,
        foreground=TEXT_DARK, font=(FONT_FAMILY, 11))
    style.map("TCombobox",
        fieldbackground=[("readonly", BG_INPUT)],
        selectbackground=[("readonly", ACCENT)])

    style.configure("TEntry",
        fieldbackground=BG_INPUT, font=(FONT_FAMILY, 11))

    style.configure("TScrollbar",
        background=BG_MAIN, troughcolor=BG_MAIN, arrowcolor=TEXT_MID)

# ─── REUSABLE WIDGETS ────────────────────────────────────────────

class SidebarBtn(tk.Button):
    """Sidebar navigation button"""
    def __init__(self, parent, text, icon="", command=None, **kw):
        super().__init__(parent,
            text=f"  {icon}  {text}",
            font=(FONT_FAMILY, 10, "bold"),
            bg=BG_SIDEBAR, fg="#AABDD4",
            activebackground="#1B6CA8", activeforeground=TEXT_WHITE,
            relief="flat", bd=0, cursor="hand2",
            anchor="w", padx=6, pady=9,
            command=command, **kw)
        self.bind("<Enter>", lambda e: self.config(bg="#162535", fg=TEXT_WHITE))
        self.bind("<Leave>", self._on_leave)

    def _on_leave(self, e):
        if not getattr(self, "_active", False):
            self.config(bg=BG_SIDEBAR, fg="#AABDD4")

    def set_active(self, active):
        self._active = active
        if active:
            self.config(bg="#1B6CA8", fg=TEXT_WHITE)
        else:
            self.config(bg=BG_SIDEBAR, fg="#AABDD4")


class Card(tk.Frame):
    """White card with shadow-like border"""
    def __init__(self, parent, **kw):
        super().__init__(parent, bg=BG_CARD,
            relief="flat", bd=1,
            highlightbackground="#D0D7E3",
            highlightthickness=1, **kw)


class KPICard(tk.Frame):
    """KPI card: icon + value + label"""
    def __init__(self, parent, icon, label, value, color=ACCENT, bg_color=None):
        super().__init__(parent, bg=bg_color or BG_CARD,
            relief="flat", bd=1,
            highlightbackground=color, highlightthickness=2,
            padx=16, pady=12)
        self.color = color

        tk.Label(self, text=icon, font=(FONT_FAMILY,22), bg=bg_color or BG_CARD,
                 fg=color).pack()
        self.val_label = tk.Label(self, text=str(value),
            font=(FONT_FAMILY,28,"bold"), bg=bg_color or BG_CARD, fg=color)
        self.val_label.pack()
        tk.Label(self, text=label, font=(FONT_FAMILY,10),
                 bg=bg_color or BG_CARD, fg=TEXT_MID).pack()

    def update_value(self, value):
        self.val_label.config(text=str(value))


class PageHeader(tk.Frame):
    """Page top header with title + optional subtitle"""
    def __init__(self, parent, title, subtitle="", color=ACCENT, **kw):
        super().__init__(parent, bg=color, pady=0, **kw)
        tk.Label(self, text=title,
            font=(FONT_FAMILY,15,"bold"), bg=color, fg=TEXT_WHITE,
            pady=12, padx=20, anchor="e").pack(fill="x")
        if subtitle:
            tk.Label(self, text=subtitle,
                font=(FONT_FAMILY,9,"italic"), bg=color, fg="#CCDDF0",
                pady=2, padx=20, anchor="e").pack(fill="x")


class ActionBtn(tk.Button):
    """Primary / secondary action button"""
    def __init__(self, parent, text, command=None, style="primary", **kw):
        colors = {
            "primary": (ACCENT,    TEXT_WHITE, "#155A9A"),
            "success": (GREEN,     TEXT_WHITE, "#185E2E"),
            "danger":  (RED,       TEXT_WHITE, "#990000"),
            "warning": (GOLD,      TEXT_WHITE, "#8B6908"),
            "neutral": ("#D0D7E3", TEXT_DARK,  "#B8BEC8"),
        }
        bg, fg, active_bg = colors.get(style, colors["primary"])
        super().__init__(parent,
            text=text, command=command,
            font=(FONT_FAMILY, 10, "bold"),
            bg=bg, fg=fg, activebackground=active_bg, activeforeground=TEXT_WHITE,
            relief="flat", bd=0, padx=14, pady=7, cursor="hand2", **kw)
        self.bind("<Enter>", lambda e: self.config(bg=active_bg))
        self.bind("<Leave>", lambda e: self.config(bg=bg))


class LabeledEntry(tk.Frame):
    """Label + Entry combo"""
    def __init__(self, parent, label, var=None, width=25, readonly=False, **kw):
        super().__init__(parent, bg=BG_CARD, **kw)
        tk.Label(self, text=label, font=F(10, True), bg=BG_CARD,
                 fg=TEXT_DARK, anchor="e").pack(anchor="e")
        state = "readonly" if readonly else "normal"
        self.entry = tk.Entry(self, textvariable=var, width=width,
                              font=F(11), justify="right",
                              bg=BG_INPUT if not readonly else BG_ROW_ALT,
                              fg=TEXT_DARK, relief="solid", bd=1,
                              state=state)
        self.entry.pack(fill="x", pady=(2,6))

    def get(self): return self.entry.get()
    def set(self, v): self.entry.delete(0,"end"); self.entry.insert(0,v)


class LabeledCombo(tk.Frame):
    """Label + Combobox combo"""
    def __init__(self, parent, label, values, var=None, width=25, **kw):
        super().__init__(parent, bg=BG_CARD, **kw)
        tk.Label(self, text=label, font=F(10, True), bg=BG_CARD,
                 fg=TEXT_DARK, anchor="e").pack(anchor="e")
        self.combo = ttk.Combobox(self, values=values, textvariable=var,
                                   width=width, font=F(11), justify="right",
                                   state="readonly")
        self.combo.pack(fill="x", pady=(2,6))

    def get(self): return self.combo.get()
    def set(self, v): self.combo.set(v)


def scrollable_frame(parent, bg=BG_MAIN):
    """Returns (outer_frame, inner_frame) — add widgets to inner_frame"""
    outer = tk.Frame(parent, bg=bg)
    canvas = tk.Canvas(outer, bg=bg, highlightthickness=0, bd=0)
    scrollbar = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
    inner = tk.Frame(canvas, bg=bg)

    inner.bind("<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0,0), window=inner, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    def _on_mousewheel(e):
        canvas.yview_scroll(int(-1*(e.delta/120)), "units")
    canvas.bind_all("<MouseWheel>", _on_mousewheel)

    return outer, inner


def make_treeview(parent, columns, headings, widths, height=15):
    """Build a styled Treeview with scrollbar"""
    frame = tk.Frame(parent, bg=BG_MAIN)
    tree = ttk.Treeview(frame, columns=columns, show="headings",
                        height=height, selectmode="browse")
    vsb = ttk.Scrollbar(frame, orient="vertical",   command=tree.yview)
    hsb = ttk.Scrollbar(frame, orient="horizontal",  command=tree.xview)
    tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

    for col, hdr, w in zip(columns, headings, widths):
        tree.heading(col, text=hdr, anchor="center")
        tree.column(col, width=w, anchor="center", minwidth=40)

    tree.tag_configure("odd",  background=BG_CARD)
    tree.tag_configure("even", background=BG_ROW_ALT)
    tree.tag_configure("low",  background=YELLOW_LT)
    tree.tag_configure("out",  background=RED_LT)
    tree.tag_configure("in",   background=GREEN_LT)

    tree.grid(row=0, column=0, sticky="nsew")
    vsb.grid(row=0, column=1, sticky="ns")
    hsb.grid(row=1, column=0, sticky="ew")
    frame.grid_rowconfigure(0, weight=1)
    frame.grid_columnconfigure(0, weight=1)

    return frame, tree


def confirm_dialog(parent, title, message):
    """Simple yes/no dialog"""
    import tkinter.messagebox as mb
    return mb.askyesno(title, message, parent=parent)


def info_dialog(parent, title, message):
    import tkinter.messagebox as mb
    mb.showinfo(title, message, parent=parent)


def error_dialog(parent, title, message):
    import tkinter.messagebox as mb
    mb.showerror(title, message, parent=parent)
