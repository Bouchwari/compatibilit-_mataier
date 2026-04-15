"""
main.py — نظام إدارة المستهلكات
الثانوية الإعدادية ألمدون — جماعة اغيل نومكون
"""
import tkinter as tk
from tkinter import ttk
import sys
import os

# Add app dir to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import database as db
from theme import *
from pages import (HomePage, DashboardPage, MovementsPage,
                   StaffPage, SuppliersPage, CertificatePage,
                   StockCardPage, StatisticsPage, ItemsManagerPage,
                   SettingsPage, YearEndPage)


class AlmadounApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("نظام إدارة المستهلكات — الثانوية الإعدادية ألمدون")
        self.geometry("1280x780")
        self.minsize(1000, 640)
        self.configure(bg=BG_DARK)

        # App icon (if available)
        try:
            self.iconbitmap(os.path.join(os.path.dirname(__file__), "assets", "icon.ico"))
        except Exception:
            pass

        apply_style(self)
        self._pages = {}
        self._nav_buttons = {}
        self._current_page = None

        self._build_layout()
        self._show_page("home")

    # ─── LAYOUT ──────────────────────────────────────────────────────
    def _build_layout(self):
        # Root: sidebar | main
        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)

        # ── SIDEBAR ───────────────────────────────────────────────────
        SBAR_W = 160
        sidebar = tk.Frame(self, bg=BG_SIDEBAR, width=SBAR_W)
        sidebar.grid(row=0, column=0, sticky="ns")
        sidebar.grid_propagate(False)
        sidebar.pack_propagate(False)

        # Logo / school name
        logo_f = tk.Frame(sidebar, bg="#0A1520", pady=10, width=SBAR_W)
        logo_f.pack(fill="x")
        logo_f.pack_propagate(False)
        tk.Label(logo_f, text="🏫", font=(FONT_FAMILY, 22), bg="#0A1520",
                 fg=TEXT_WHITE, width=SBAR_W).pack()
        tk.Label(logo_f, text="إدارة\nالمستهلكات",
                 font=(FONT_FAMILY, 10, "bold"), bg="#0A1520",
                 fg=TEXT_WHITE, wraplength=SBAR_W - 10, justify="center").pack()
        tk.Label(logo_f, text="ألمدون",
                 font=(FONT_FAMILY, 8, "italic"), bg="#0A1520",
                 fg="#7090B0", wraplength=SBAR_W - 10).pack()

        tk.Frame(sidebar, bg="#1F3347", height=1).pack(fill="x", pady=3)

        # Navigation items — inside a scrollable canvas
        nav_items = [
            ("🏠", "الرئيسية",          "home"),
            ("📊", "لوحة التحكم",       "dashboard"),
            ("📈", "الإحصائيات",        "statistics"),
            ("──────────","",           "sep1"),
            ("📁", "مكتبيات",           "cat_مكتبيات"),
            ("📚", "أدوات التدريس",     "cat_ادوات التدريس"),
            ("🧹", "أدوات النظافة",     "cat_ادوات النظافة"),
            ("🔧", "الصيانة",           "cat_الصيانة"),
            ("📦", "مختلفات",           "cat_مختلفات"),
            ("──────────","",           "sep2"),
            ("🗒️", "إدارة المواد",      "items"),
            ("👥", "الأساتذة",          "staff"),
            ("🏪", "الموردون",          "suppliers"),
            ("📋", "شهادة التسلم",      "certificate"),
            ("🗃️", "بطاقة المخزون",    "stockcard"),
            ("──────────","",           "sep3"),
            ("⚙️", "إعدادات المؤسسة",  "settings"),
            ("📅", "نهاية / بداية سنة", "yearend"),
        ]

        # Scrollable canvas for nav
        nav_canvas = tk.Canvas(sidebar, bg=BG_SIDEBAR,
                               highlightthickness=0, bd=0,
                               width=SBAR_W)
        nav_canvas.pack(fill="both", expand=True)

        nav_scroll = tk.Scrollbar(sidebar, orient="vertical",
                                  command=nav_canvas.yview)
        # Only show scrollbar when needed — keep sidebar clean
        nav_canvas.configure(yscrollcommand=nav_scroll.set)

        nav_inner = tk.Frame(nav_canvas, bg=BG_SIDEBAR)
        nav_window = nav_canvas.create_window((0, 0), window=nav_inner,
                                              anchor="nw")

        def _on_nav_configure(e):
            nav_canvas.configure(scrollregion=nav_canvas.bbox("all"))
            nav_canvas.itemconfig(nav_window,
                                  width=nav_canvas.winfo_width())
        nav_inner.bind("<Configure>", _on_nav_configure)
        nav_canvas.bind("<Configure>",
            lambda e: nav_canvas.itemconfig(nav_window,
                                            width=e.width))

        def _on_mousewheel(e):
            nav_canvas.yview_scroll(int(-1*(e.delta/120)), "units")
        nav_canvas.bind_all("<MouseWheel>", _on_mousewheel)

        for icon, label, key in nav_items:
            if key.startswith("sep"):
                tk.Frame(nav_inner, bg="#1F3347", height=1).pack(
                    fill="x", pady=3)
                continue
            btn = SidebarBtn(nav_inner, label, icon=icon,
                             command=lambda k=key: self._show_page(k))
            btn.pack(fill="x")
            self._nav_buttons[key] = btn

        # Version label at bottom of canvas
        tk.Label(nav_inner, text="v1.0 — 2025/2026",
                 font=(FONT_FAMILY, 8), bg=BG_SIDEBAR,
                 fg="#4A6A8A", pady=10).pack(fill="x")

        # ── MAIN AREA ─────────────────────────────────────────────────
        self.main_frame = tk.Frame(self, bg=BG_MAIN)
        self.main_frame.grid(row=0, column=1, sticky="nsew")
        self.main_frame.rowconfigure(0, weight=1)
        self.main_frame.columnconfigure(0, weight=1)

    # ─── PAGE NAVIGATION ─────────────────────────────────────────────
    def _get_or_create_page(self, key):
        if key not in self._pages:
            if key == "home":
                page = HomePage(self.main_frame)
            elif key == "dashboard":
                page = DashboardPage(self.main_frame)
            elif key == "statistics":
                page = StatisticsPage(self.main_frame)
            elif key == "staff":
                page = StaffPage(self.main_frame)
            elif key == "suppliers":
                page = SuppliersPage(self.main_frame)
            elif key == "certificate":
                page = CertificatePage(self.main_frame)
            elif key == "stockcard":
                page = StockCardPage(self.main_frame)
            elif key == "items":
                page = ItemsManagerPage(self.main_frame)
            elif key == "settings":
                page = SettingsPage(self.main_frame)
            elif key == "yearend":
                page = YearEndPage(self.main_frame)
            elif key.startswith("cat_"):
                cat_name = key[4:]
                page = MovementsPage(self.main_frame, category=cat_name)
            else:
                page = HomePage(self.main_frame)

            page.grid(row=0, column=0, sticky="nsew")
            self._pages[key] = page

        return self._pages[key]

    def _show_page(self, key):
        # Deactivate all nav buttons
        for k, btn in self._nav_buttons.items():
            btn.set_active(k == key)

        # Hide current page
        if self._current_page:
            self._current_page.grid_remove()

        # Show new page
        page = self._get_or_create_page(key)
        page.grid()

        # Refresh page data
        if hasattr(page, "refresh"):
            page.refresh()

        self._current_page = page

    def run(self):
        self.mainloop()


# ─── SPLASH SCREEN ───────────────────────────────────────────────
def show_splash():
    splash = tk.Tk()
    splash.overrideredirect(True)
    w, h = 480, 280
    x = (splash.winfo_screenwidth()  - w) // 2
    y = (splash.winfo_screenheight() - h) // 2
    splash.geometry(f"{w}x{h}+{x}+{y}")
    splash.configure(bg=BG_DARK)

    tk.Label(splash, text="🏫", font=(FONT_FAMILY,48), bg=BG_DARK, fg=TEXT_WHITE).pack(pady=(30,10))
    tk.Label(splash, text="نظام إدارة المستهلكات",
             font=(FONT_FAMILY,16,"bold"), bg=BG_DARK, fg=TEXT_WHITE).pack()
    tk.Label(splash, text="الثانوية الإعدادية ألمدون",
             font=(FONT_FAMILY,12), bg=BG_DARK, fg="#7090B0").pack(pady=4)

    prog = ttk.Progressbar(splash, length=300, mode="determinate")
    prog.pack(pady=20)

    status = tk.Label(splash, text="جاري التحميل...",
                      font=(FONT_FAMILY,9), bg=BG_DARK, fg="#5A7A9A")
    status.pack()

    msgs = ["تهيئة قاعدة البيانات...","تحميل البيانات...","جاري الإعداد...","اكتمل التحميل ✅"]
    def step(i=0):
        if i < len(msgs):
            prog["value"] = (i+1) * 25
            status.config(text=msgs[i])
            splash.after(300, step, i+1)
        else:
            splash.after(200, splash.destroy)
    step()
    splash.mainloop()


def _show_setup_wizard():
    """Modal setup dialog shown on first run to collect school info."""
    from tkinter import filedialog, messagebox as _mb

    win = tk.Toplevel()
    win.title("إعداد المؤسسة — أول تشغيل")
    win.configure(bg=BG_DARK)
    win.resizable(False, False)
    w, h = 540, 460
    x = (win.winfo_screenwidth()  - w) // 2
    y = (win.winfo_screenheight() - h) // 2
    win.geometry(f"{w}x{h}+{x}+{y}")
    win.grab_set()

    tk.Label(win, text="🏫  مرحباً! إعداد المؤسسة",
             font=(FONT_FAMILY, 15, "bold"), bg=BG_DARK, fg=TEXT_WHITE,
             pady=16).pack(fill="x")
    tk.Label(win,
             text="يرجى ملء المعلومات التالية — ستظهر في الوثائق الرسمية",
             font=(FONT_FAMILY, 10), bg=BG_DARK, fg="#7090B0").pack()

    frm = tk.Frame(win, bg=BG_CARD, padx=20, pady=16)
    frm.pack(fill="x", padx=20, pady=12)

    fields = [
        ("اسم المؤسسة *",       "school_name"),
        ("الأكاديمية",           "academy"),
        ("المديرية الإقليمية",   "delegation"),
        ("العنوان / المدينة",    "address"),
    ]
    vars_ = {}
    logo_var = tk.StringVar()

    for lbl, key in fields:
        r = tk.Frame(frm, bg=BG_CARD); r.pack(fill="x", pady=4)
        tk.Label(r, text=lbl + ":", font=(FONT_FAMILY, 10, "bold"),
                 bg=BG_CARD, fg=TEXT_DARK, width=22, anchor="e").pack(side="right")
        v = tk.StringVar()
        vars_[key] = v
        tk.Entry(r, textvariable=v, font=(FONT_FAMILY, 11), justify="right",
                 bg=BG_INPUT, relief="solid", bd=1).pack(
                 side="right", fill="x", expand=True, padx=(0, 6))

    # Logo row
    lr = tk.Frame(frm, bg=BG_CARD); lr.pack(fill="x", pady=4)
    tk.Label(lr, text="الشعار (اختياري):", font=(FONT_FAMILY, 10, "bold"),
             bg=BG_CARD, fg=TEXT_DARK, width=22, anchor="e").pack(side="right")
    tk.Entry(lr, textvariable=logo_var, font=(FONT_FAMILY, 9), justify="right",
             bg=BG_INPUT, relief="solid", bd=1, state="readonly",
             width=28).pack(side="right", padx=(4, 6))

    def pick_logo():
        p = filedialog.askopenfilename(
            parent=win, title="اختر شعار المؤسسة",
            filetypes=[("Images", "*.png *.jpg *.jpeg *.bmp")])
        if p:
            logo_var.set(p)
    tk.Button(lr, text="📂 اختيار", font=(FONT_FAMILY, 9), relief="flat",
              bg=ACCENT, fg="white", cursor="hand2",
              command=pick_logo).pack(side="right")

    def do_save():
        name = vars_["school_name"].get().strip()
        if not name:
            _mb.showwarning("تنبيه", "يرجى إدخال اسم المؤسسة", parent=win)
            return
        db.save_settings(
            name,
            vars_["academy"].get().strip(),
            vars_["delegation"].get().strip(),
            vars_["address"].get().strip(),
            logo_var.get().strip()
        )
        win.destroy()

    def do_skip():
        win.destroy()

    btn_row = tk.Frame(win, bg=BG_DARK); btn_row.pack(fill="x", padx=20, pady=10)
    tk.Button(btn_row, text="✅  حفظ والمتابعة",
              font=(FONT_FAMILY, 12, "bold"),
              bg="#06D6A0", fg="white", relief="flat", cursor="hand2",
              pady=10, command=do_save).pack(side="right", fill="x", expand=True, padx=(4,0))
    tk.Button(btn_row, text="تخطي",
              font=(FONT_FAMILY, 10),
              bg="#4A6080", fg="white", relief="flat", cursor="hand2",
              pady=10, command=do_skip).pack(side="right", padx=(0,4))

    win.wait_window()


if __name__ == "__main__":
    # Initialize DB (also runs migrations)
    db.init_db()
    # First-run setup wizard — needs a hidden root to show Toplevel
    if db.is_first_run():
        _root = tk.Tk()
        _root.withdraw()          # hide the temporary root
        _show_setup_wizard()
        _root.destroy()
    # Show splash
    show_splash()
    # Launch app
    app = AlmadounApp()
    app.run()

