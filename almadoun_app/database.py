"""
database.py — SQLite schema + all CRUD operations
الثانوية الإعدادية ألمدون — نظام إدارة المستهلكات
"""
import sqlite3
import os
from datetime import datetime

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "almadoun.db")

# ─── ALL ITEMS ──────────────────────────────────────────────────────
CATEGORIES = {
    "مكتبيات": [
        ("أوراق A4 ليزر أبيض 80g","رزمة 500"),("ورق مقوى للتغليف A4","رزمة 100"),
        ("ورق شفاف للتجليد A4","رزمة 100"),("ورق A4 لاصق","رزمة 500"),
        ("ورق A4 ملون أخضر","رزمة 500"),("ورق A4 ملون أصفر","رزمة 500"),
        ("ورق A4 ملون وردي","رزمة 500"),("أقلام جافة زرقاء","رزمة 50"),
        ("أقلام ماركر سبورة","قطعة"),("أقلام فلوريسان ملونة","قطعة"),
        ("أقلام ماركر دائمة","قطعة"),("أقلام رصاص","علبة"),
        ("قلم لاصق Glue Stick","قطعة"),("أعواد تجليد مختلفة","علبة 25"),
        ("علبة أرشيف كرتون 38x29x12","قطعة"),("علبة أرشيف كرتون 34x26x8","قطعة"),
        ("ملفات Bull صفراء","رزمة 250"),("ملفات مقوى ملونة","رزمة 100"),
        ("ملف كرتون بأرجل","قطعة"),("ملف بلاستيك شفاف","قطعة"),
        ("دباسة ورق","قطعة"),("علبة دباسيس 24/6","علبة 10"),
        ("مشابك ورق","علبة"),("دبابيس بيضاء","علبة 50"),
        ("أظرف صغيرة 18x12","قطعة"),("أظرف متوسطة 24x19","قطعة"),
        ("أظرف كبيرة 36x28","قطعة"),("سجل الصادر","قطعة"),
        ("سجل الوارد","قطعة"),("دفتر حلزوني A4","قطعة"),
        ("دفاتر عادية","قطعة"),("لاصق شريطي شفاف","بكرة"),
        ("مقص","قطعة"),("مسطرة 30cm","قطعة"),
        ("خاتم تاريخ فرنسي","قطعة"),("خاتم تاريخ عربي","قطعة"),
        ("وسادة حبر للختم","قطعة"),("أسطوانات CD فارغة","علبة 50"),
        ("أسطوانات DVD فارغة","قطعة"),("علم المغرب 90x150","قطعة"),
    ],
    "خراطيش الطباعة": [
        ("خرطوشة Lexmark E120","قطعة"),("خرطوشة Lexmark E240","قطعة"),
        ("خرطوشة Lexmark E260","قطعة"),("خرطوشة Lexmark B2236DW","قطعة"),
        ("خرطوشة Lexmark B2338","قطعة"),("خرطوشة Lexmark MX511","قطعة"),
        ("خرطوشة Canon LBP2900","قطعة"),("خرطوشة Canon MF3010","قطعة"),
        ("خرطوشة Canon MF443DW","قطعة"),("خرطوشة Canon iR1435","قطعة"),
        ("خرطوشة Canon iR2202","قطعة"),("خرطوشة Canon iR2204","قطعة"),
        ("خرطوشة Canon iR2206","قطعة"),("خرطوشة Canon LBP226DW","قطعة"),
        ("خرطوشة Canon 2420","قطعة"),("خرطوشة Canon C-EXV 14","قطعة"),
        ("خرطوشة Canon C-EXV 42","قطعة"),("خرطوشة Canon PC D320","قطعة"),
        ("خرطوشة HP LJ P1005","قطعة"),("خرطوشة HP 85A P1102","قطعة"),
        ("خرطوشة HP LJ 1018","قطعة"),("خرطوشة HP LJ 1022","قطعة"),
        ("خرطوشة HP LJ 1320","قطعة"),("خرطوشة HP LJ 17A","قطعة"),
        ("خرطوشة HP 35A","قطعة"),("خرطوشة HP LJ 2300","قطعة"),
        ("خرطوشة HP MFP M26a","قطعة"),("خرطوشة HP MFP M130A","قطعة"),
        ("خرطوشة HP MFP 130FN","قطعة"),("خرطوشة Samsung ML-1660","قطعة"),
        ("خرطوشة Samsung M2070","قطعة"),("خرطوشة Samsung D3470B","قطعة"),
        ("خرطوشة Samsung M3820","قطعة"),("خرطوشة Konica Bizhub 185","قطعة"),
        ("خرطوشة Konica Bizhub 215","قطعة"),("خرطوشة Konica Bizhub 250","قطعة"),
        ("خرطوشة Konica Bizhub 282","قطعة"),("خرطوشة Konica Bizhub 283","قطعة"),
        ("خرطوشة Konica Bizhub 364E","قطعة"),("خرطوشة Konica Bizhub 423","قطعة"),
        ("خرطوشة Toshiba Studio 163","قطعة"),("خرطوشة Sharp AR-5516","قطعة"),
        ("خرطوشة Ricoh SPF 201","قطعة"),("خرطوشة Kyocera KM-2035","قطعة"),
    ],
    "ادوات التدريس": [
        ("طباشير أبيض بدون غبار","علبة 100"),("طباشير ملون بدون غبار","علبة 100"),
        ("ممسحة سبورة طباشير","قطعة"),("ممسحة سبورة مغناطيسية","قطعة"),
        ("قلم ماركر قابل للمسح","قطعة"),("حشو ماركر أسود","علبة"),
        ("حشو ماركر أحمر","علبة"),("حشو ماركر أخضر","علبة"),
        ("حشو ماركر أزرق","علبة"),("مسطرة كبيرة 100cm","قطعة"),
        ("فرجار بشفط للسبورة","قطعة"),("فرجار للسبورة الخشبية","قطعة"),
        ("منقلة للسبورة","قطعة"),("مثلث قياس 60cm","قطعة"),
        ("خرائط جغرافية","قطعة"),("كتب مرجعية","قطعة"),
        ("أقراص CD/DVD تعليمية","قطعة"),
    ],
    "ادوات النظافة": [
        ("مسحوق التنظيف","كيس"),("صابون سائل","قارورة"),
        ("ماء جافيل","قارورة"),("منظف الأرضيات","قارورة"),
        ("منظف الزجاج","قارورة"),("مكنسة يدوية","قطعة"),
        ("مكنسة عصا","قطعة"),("لفة جفاف","قطعة"),
        ("إسفنجة تنظيف","قطعة"),("قفازات مطاطية","زوج"),
        ("كيس القمامة","علبة"),("ورق التواليت","رزمة"),
        ("مناديل ورقية","علبة"),("دلو","قطعة"),
        ("بالوعة مطاطية","قطعة"),("سطل","قطعة"),
        ("عود مسح","قطعة"),
    ],
    "الصيانة": [
        ("مسامير متنوعة","علبة"),("براغي متنوعة","علبة"),
        ("دهان أبيض","قدح"),("فرشاة دهان","قطعة"),
        ("رول دهان","قطعة"),("تينر","قارورة"),
        ("لمبات إضاءة","قطعة"),("قواطع كهربائية","قطعة"),
        ("أسلاك كهربائية","متر"),("مقابس كهربائية","قطعة"),
        ("شريط لاصق عازل","بكرة"),("صنبور ماء","قطعة"),
        ("سيلكون لاصق","قطعة"),("مفتاح ربط","قطعة"),
        ("مطرقة","قطعة"),("زجاج نافذة","قطعة"),
        ("قفل باب","قطعة"),("مفصلة باب","قطعة"),
        ("مواد سباكة متنوعة","قطعة"),
    ],
    "مختلفات": [
        ("بطاريات AA","قطعة"),("بطاريات AAA","قطعة"),
        ("قهوة","علبة"),("شاي","علبة"),
        ("سكر","كيلو"),("كؤوس ورقية","علبة"),
        ("كؤوس بلاستيك","علبة"),("ملاعق بلاستيك","علبة"),
        ("مياه معدنية","قارورة"),("علب حفظ","قطعة"),
        ("لافتات","قطعة"),("إطارات صور","قطعة"),
        ("شمع","قطعة"),("مبيد حشري","علبة"),
        ("هواء مضغوط للحاسوب","علبة"),
    ],
}

def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn

def init_db():
    conn = get_conn()
    c = conn.cursor()

    c.executescript("""
    CREATE TABLE IF NOT EXISTS items (
        id       INTEGER PRIMARY KEY AUTOINCREMENT,
        name     TEXT NOT NULL UNIQUE,
        unit     TEXT,
        category TEXT
    );

    CREATE TABLE IF NOT EXISTS movements (
        id          INTEGER PRIMARY KEY AUTOINCREMENT,
        item_name   TEXT NOT NULL,
        category    TEXT,
        date        TEXT,
        type        TEXT CHECK(type IN ('دخول','خروج')),
        quantity    REAL DEFAULT 0,
        beneficiary TEXT,
        notes       TEXT,
        created_at  TEXT DEFAULT (datetime('now'))
    );

    CREATE TABLE IF NOT EXISTS staff (
        id          INTEGER PRIMARY KEY AUTOINCREMENT,
        name        TEXT NOT NULL,
        position    TEXT,
        subject     TEXT,
        hire_number TEXT,
        email       TEXT,
        notes       TEXT
    );

    CREATE TABLE IF NOT EXISTS settings (
        id           INTEGER PRIMARY KEY CHECK(id=1),
        school_name  TEXT DEFAULT '',
        academy      TEXT DEFAULT '',
        delegation   TEXT DEFAULT '',
        address      TEXT DEFAULT '',
        logo_path    TEXT DEFAULT ''
    );

    CREATE TABLE IF NOT EXISTS year_closings (
        id           INTEGER PRIMARY KEY AUTOINCREMENT,
        year_label   TEXT NOT NULL UNIQUE,
        closed_at    TEXT DEFAULT (datetime('now')),
        snapshot_json TEXT
    );

    CREATE TABLE IF NOT EXISTS suppliers (
        id           INTEGER PRIMARY KEY AUTOINCREMENT,
        name         TEXT NOT NULL,
        supply_type  TEXT,
        phone        TEXT,
        address      TEXT,
        last_invoice TEXT,
        last_date    TEXT,
        notes        TEXT
    );

    CREATE TABLE IF NOT EXISTS certificates (
        id          INTEGER PRIMARY KEY AUTOINCREMENT,
        cert_number INTEGER,
        recipient   TEXT,
        position    TEXT,
        date        TEXT,
        items_json  TEXT,
        created_at  TEXT DEFAULT (datetime('now'))
    );
    """)

    # Migrate: rename phone -> hire_number in staff (safe, skipped if already done)
    cols = [r[1] for r in c.execute("PRAGMA table_info(staff)").fetchall()]
    if "phone" in cols and "hire_number" not in cols:
        c.execute("ALTER TABLE staff RENAME COLUMN phone TO hire_number")

    # Seed settings row (singleton)
    c.execute("INSERT OR IGNORE INTO settings (id) VALUES (1)")

    # Seed items if empty
    if c.execute("SELECT COUNT(*) FROM items").fetchone()[0] == 0:
        for cat, items in CATEGORIES.items():
            for name, unit in items:
                c.execute("INSERT OR IGNORE INTO items (name,unit,category) VALUES (?,?,?)",
                          (name, unit, cat))

    conn.commit()
    conn.close()

# ─── ITEMS ──────────────────────────────────────────────────────────
def get_all_items():
    with get_conn() as conn:
        return conn.execute("SELECT * FROM items ORDER BY category, name").fetchall()

def get_items_by_category(category):
    with get_conn() as conn:
        return conn.execute(
            "SELECT * FROM items WHERE category=? ORDER BY name", (category,)
        ).fetchall()

def get_item_unit(item_name):
    with get_conn() as conn:
        row = conn.execute("SELECT unit FROM items WHERE name=?", (item_name,)).fetchone()
        return row["unit"] if row else ""

def add_item(name, unit, category):
    with get_conn() as conn:
        conn.execute("INSERT INTO items (name,unit,category) VALUES (?,?,?)", (name,unit,category))
        conn.commit()

def delete_item(item_id):
    with get_conn() as conn:
        conn.execute("DELETE FROM items WHERE id=?", (item_id,))
        conn.commit()

def update_item(item_id, name, unit, category):
    with get_conn() as conn:
        conn.execute(
            "UPDATE items SET name=?, unit=?, category=? WHERE id=?",
            (name, unit, category, item_id)
        )
        conn.commit()

def item_has_movements(item_name):
    """Returns count of movements for an item (to warn before delete)."""
    with get_conn() as conn:
        return conn.execute(
            "SELECT COUNT(*) FROM movements WHERE item_name=?", (item_name,)
        ).fetchone()[0]

# ─── MOVEMENTS ──────────────────────────────────────────────────────
def add_movement(item_name, category, date, mov_type, quantity, beneficiary, notes):
    with get_conn() as conn:
        conn.execute("""
            INSERT INTO movements (item_name,category,date,type,quantity,beneficiary,notes)
            VALUES (?,?,?,?,?,?,?)
        """, (item_name, category, date, mov_type, quantity, beneficiary, notes))
        conn.commit()

def get_movements(category=None, item_name=None, limit=None):
    with get_conn() as conn:
        q = "SELECT * FROM movements WHERE 1=1"
        params = []
        if category:
            q += " AND category=?"; params.append(category)
        if item_name:
            q += " AND item_name=?"; params.append(item_name)
        q += " ORDER BY date DESC, id DESC"
        if limit:
            q += f" LIMIT {limit}"
        return conn.execute(q, params).fetchall()

def delete_movement(mov_id):
    with get_conn() as conn:
        conn.execute("DELETE FROM movements WHERE id=?", (mov_id,))
        conn.commit()

def get_item_balance(item_name):
    with get_conn() as conn:
        in_ = conn.execute(
            "SELECT COALESCE(SUM(quantity),0) FROM movements WHERE item_name=? AND type='دخول'",
            (item_name,)
        ).fetchone()[0]
        out_ = conn.execute(
            "SELECT COALESCE(SUM(quantity),0) FROM movements WHERE item_name=? AND type='خروج'",
            (item_name,)
        ).fetchone()[0]
        return in_, out_, in_ - out_

def get_dashboard_data():
    """Returns list of (item_name, unit, category, total_in, total_out, balance, last_date)"""
    with get_conn() as conn:
        rows = conn.execute("""
            SELECT
                i.name, i.unit, i.category,
                COALESCE(SUM(CASE WHEN m.type='دخول' THEN m.quantity ELSE 0 END),0) as total_in,
                COALESCE(SUM(CASE WHEN m.type='خروج' THEN m.quantity ELSE 0 END),0) as total_out,
                COALESCE(SUM(CASE WHEN m.type='دخول' THEN m.quantity ELSE 0 END),0)
                    - COALESCE(SUM(CASE WHEN m.type='خروج' THEN m.quantity ELSE 0 END),0) as balance,
                MAX(m.date) as last_date
            FROM items i
            LEFT JOIN movements m ON m.item_name = i.name
            GROUP BY i.name, i.unit, i.category
            ORDER BY i.category, i.name
        """).fetchall()
        return rows

def get_category_stats():
    """Returns per-category totals for charts"""
    with get_conn() as conn:
        return conn.execute("""
            SELECT
                i.category,
                COALESCE(SUM(CASE WHEN m.type='دخول' THEN m.quantity ELSE 0 END),0) as total_in,
                COALESCE(SUM(CASE WHEN m.type='خروج' THEN m.quantity ELSE 0 END),0) as total_out,
                COUNT(DISTINCT m.id) as operations
            FROM items i
            LEFT JOIN movements m ON m.item_name = i.name
            GROUP BY i.category
        """).fetchall()

def get_monthly_stats(year=None):
    if not year:
        year = datetime.now().year
    with get_conn() as conn:
        return conn.execute("""
            SELECT
                CAST(strftime('%m', date) AS INTEGER) as month,
                COALESCE(SUM(CASE WHEN type='دخول' THEN quantity ELSE 0 END),0) as total_in,
                COALESCE(SUM(CASE WHEN type='خروج' THEN quantity ELSE 0 END),0) as total_out,
                COUNT(*) as operations
            FROM movements
            WHERE strftime('%Y', date) = ?
            GROUP BY month
            ORDER BY month
        """, (str(year),)).fetchall()

def get_kpi_summary():
    data = get_dashboard_data()
    total = len(data)
    available = sum(1 for r in data if r["balance"] > 5)
    low       = sum(1 for r in data if 0 < r["balance"] <= 5)
    out_      = sum(1 for r in data if r["balance"] <= 0)
    ops       = 0
    with get_conn() as conn:
        ops = conn.execute("SELECT COUNT(*) FROM movements").fetchone()[0]
    return {"total": total, "available": available, "low": low, "out": out_, "operations": ops}

def get_top_consumed_items(n=10):
    """Top N items by total outflow quantity."""
    with get_conn() as conn:
        return conn.execute("""
            SELECT item_name, category,
                SUM(CASE WHEN type='دخول' THEN quantity ELSE 0 END) as total_in,
                SUM(CASE WHEN type='خروج' THEN quantity ELSE 0 END) as total_out,
                COUNT(*) as ops
            FROM movements
            GROUP BY item_name
            ORDER BY total_out DESC LIMIT ?
        """, (n,)).fetchall()

def get_top_beneficiaries(n=10):
    """Top N people/entities receiving items (خروج movements)."""
    with get_conn() as conn:
        return conn.execute("""
            SELECT beneficiary, COUNT(*) as ops,
                SUM(quantity) as total_qty,
                COUNT(DISTINCT item_name) as distinct_items
            FROM movements
            WHERE type='خروج'
              AND beneficiary IS NOT NULL AND beneficiary != ''
            GROUP BY beneficiary
            ORDER BY total_qty DESC LIMIT ?
        """, (n,)).fetchall()

def get_stock_alerts():
    """Items with balance <= 0 (out) or <= 5 (low), ordered by severity."""
    with get_conn() as conn:
        return conn.execute("""
            SELECT i.name, i.category, i.unit,
                COALESCE(SUM(CASE WHEN m.type='دخول' THEN m.quantity ELSE 0 END),0)
                - COALESCE(SUM(CASE WHEN m.type='خروج' THEN m.quantity ELSE 0 END),0) as balance,
                MAX(m.date) as last_date
            FROM items i
            LEFT JOIN movements m ON m.item_name = i.name
            GROUP BY i.name, i.category, i.unit
            HAVING balance <= 5
            ORDER BY balance ASC, i.category
        """).fetchall()

def get_movement_trend(months=12):
    """Monthly movement totals for the last N months."""
    with get_conn() as conn:
        return conn.execute("""
            SELECT strftime('%Y-%m', date) as ym,
                SUM(CASE WHEN type='دخول' THEN quantity ELSE 0 END) as total_in,
                SUM(CASE WHEN type='خروج' THEN quantity ELSE 0 END) as total_out,
                COUNT(*) as ops
            FROM movements
            WHERE date >= date('now', ? || ' months')
            GROUP BY ym
            ORDER BY ym
        """, (f"-{months}",)).fetchall()


# ─── STAFF ──────────────────────────────────────────────────────────
def get_all_staff():
    with get_conn() as conn:
        return conn.execute("SELECT * FROM staff ORDER BY name").fetchall()

def get_staff_names():
    with get_conn() as conn:
        return [r["name"] for r in conn.execute("SELECT name FROM staff ORDER BY name").fetchall()]

def get_staff_stats():
    """Returns per-teacher withdrawal totals for statistics."""
    with get_conn() as conn:
        return conn.execute("""
            SELECT
                beneficiary as name,
                COUNT(*) as operations,
                COALESCE(SUM(quantity), 0) as total_qty,
                COUNT(DISTINCT category) as categories_count
            FROM movements
            WHERE type='خروج'
              AND beneficiary IS NOT NULL
              AND beneficiary != ''
            GROUP BY beneficiary
            ORDER BY total_qty DESC
        """).fetchall()

def add_staff(name, position, subject, hire_number, email, notes):
    with get_conn() as conn:
        conn.execute("""
            INSERT INTO staff (name,position,subject,hire_number,email,notes)
            VALUES (?,?,?,?,?,?)
        """, (name, position, subject, hire_number, email, notes))
        conn.commit()

def update_staff(staff_id, name, position, subject, hire_number, email, notes):
    with get_conn() as conn:
        conn.execute("""
            UPDATE staff SET name=?,position=?,subject=?,hire_number=?,email=?,notes=?
            WHERE id=?
        """, (name, position, subject, hire_number, email, notes, staff_id))
        conn.commit()

def delete_staff(staff_id):
    with get_conn() as conn:
        conn.execute("DELETE FROM staff WHERE id=?", (staff_id,))
        conn.commit()

# ─── SUPPLIERS ──────────────────────────────────────────────────────
def get_all_suppliers():
    with get_conn() as conn:
        return conn.execute("SELECT * FROM suppliers ORDER BY name").fetchall()

def get_supplier_names():
    with get_conn() as conn:
        return [r["name"] for r in conn.execute("SELECT name FROM suppliers ORDER BY name").fetchall()]

def add_supplier(name, supply_type, phone, address, last_invoice, last_date, notes):
    with get_conn() as conn:
        conn.execute("""
            INSERT INTO suppliers (name,supply_type,phone,address,last_invoice,last_date,notes)
            VALUES (?,?,?,?,?,?,?)
        """, (name, supply_type, phone, address, last_invoice, last_date, notes))
        conn.commit()

def update_supplier(sup_id, name, supply_type, phone, address, last_invoice, last_date, notes):
    with get_conn() as conn:
        conn.execute("""
            UPDATE suppliers SET name=?,supply_type=?,phone=?,address=?,
            last_invoice=?,last_date=?,notes=? WHERE id=?
        """, (name, supply_type, phone, address, last_invoice, last_date, notes, sup_id))
        conn.commit()

def delete_supplier(sup_id):
    with get_conn() as conn:
        conn.execute("DELETE FROM suppliers WHERE id=?", (sup_id,))
        conn.commit()

# ─── CERTIFICATES ───────────────────────────────────────────────────
def get_next_cert_number():
    with get_conn() as conn:
        row = conn.execute("SELECT MAX(cert_number) FROM certificates").fetchone()[0]
        return (row or 0) + 1

def save_certificate(cert_number, recipient, position, date, items_json):
    with get_conn() as conn:
        conn.execute("""
            INSERT INTO certificates (cert_number,recipient,position,date,items_json)
            VALUES (?,?,?,?,?)
        """, (cert_number, recipient, position, date, items_json))
        conn.commit()

def get_certificates():
    with get_conn() as conn:
        return conn.execute(
            "SELECT * FROM certificates ORDER BY created_at DESC"
        ).fetchall()

# Init on import
init_db()

# ─── MOVEMENTS ──────────────────────────────────────────────────────
def delete_movement(movement_id):
    """Permanently delete a single movement record."""
    with get_conn() as conn:
        conn.execute("DELETE FROM movements WHERE id=?", (movement_id,))
        conn.commit()

# ─── SETTINGS ───────────────────────────────────────────────────────
def get_settings():
    """Return the singleton settings row as a dict."""
    with get_conn() as conn:
        row = conn.execute("SELECT * FROM settings WHERE id=1").fetchone()
        if row:
            return dict(row)
        return {"school_name":"","academy":"","delegation":"",
                "address":"","logo_path":""}

def save_settings(school_name, academy, delegation, address, logo_path):
    with get_conn() as conn:
        conn.execute("""
            INSERT INTO settings (id,school_name,academy,delegation,address,logo_path)
            VALUES (1,?,?,?,?,?)
            ON CONFLICT(id) DO UPDATE SET
                school_name=excluded.school_name,
                academy=excluded.academy,
                delegation=excluded.delegation,
                address=excluded.address,
                logo_path=excluded.logo_path
        """, (school_name, academy, delegation, address, logo_path))
        conn.commit()

def is_first_run():
    """True if settings have never been filled."""
    s = get_settings()
    return not s.get("school_name", "").strip()

# ─── YEAR END ────────────────────────────────────────────────────────
def close_year(year_label):
    """
    Snapshot current balances and save to year_closings.
    Returns (export_path, item_count).
    """
    import json
    data = get_dashboard_data()
    snapshot = [
        {"name": r["name"], "category": r["category"],
         "unit": r["unit"], "balance": r["balance"]}
        for r in data
    ]
    with get_conn() as conn:
        conn.execute("""
            INSERT INTO year_closings (year_label, snapshot_json)
            VALUES (?,?)
            ON CONFLICT(year_label) DO UPDATE SET
                snapshot_json=excluded.snapshot_json,
                closed_at=datetime('now')
        """, (year_label, json.dumps(snapshot, ensure_ascii=False)))
        conn.commit()
    return snapshot

def open_new_year(year_label, from_snapshot=None):
    """
    Create initial 'دخول' entries for each item using their last balance.
    from_snapshot: list of dicts from close_year(); if None, uses current balances.
    Returns count of entries created.
    """
    if from_snapshot is None:
        from_snapshot = [
            {"name": r["name"], "category": r["category"],
             "unit": r["unit"], "balance": r["balance"]}
            for r in get_dashboard_data()
        ]
    today = datetime.now().strftime("%Y-%m-%d")
    count = 0
    with get_conn() as conn:
        for item in from_snapshot:
            bal = float(item.get("balance", 0))
            if bal <= 0:
                continue
            conn.execute("""
                INSERT INTO movements
                    (item_name, category, date, type, quantity, beneficiary, notes)
                VALUES (?,?,?,?,?,?,?)
            """, (item["name"], item["category"], today,
                   "دخول", bal, "",
                   f"رصيد مُرحَّل من سنة {year_label}"))
            count += 1
        conn.commit()
    return count

def get_year_closings():
    with get_conn() as conn:
        return conn.execute(
            "SELECT id, year_label, closed_at FROM year_closings ORDER BY closed_at DESC"
        ).fetchall()
