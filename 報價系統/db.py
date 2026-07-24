# -*- coding: utf-8 -*-
"""
SQLite 資料層 ── 單機、單一檔案 補習班.db，免安裝伺服器。

資料表
    students      學生名冊
    courses       課程／價目表（預設學費、教材費）
    quotes        報價主檔（含金額小計、繳費狀態）
    quote_items   報價明細列（每一堂課程一列）

所有金額以整數（新台幣元）儲存。報價成立時會把學生姓名等「快照」寫進
quotes / quote_items，之後即使改了名冊，歷史報價單仍保留當時內容。
"""
from __future__ import annotations

import sqlite3
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path

DB_PATH = Path(__file__).parent / "補習班.db"

SCHEMA = """
CREATE TABLE IF NOT EXISTS students (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    name        TEXT NOT NULL,
    grade       TEXT,
    parent_name TEXT,
    phone       TEXT,
    email       TEXT,
    address     TEXT,
    note        TEXT,
    created_at  TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS courses (
    id               INTEGER PRIMARY KEY AUTOINCREMENT,
    name             TEXT NOT NULL,
    default_tuition  INTEGER NOT NULL DEFAULT 0,
    default_material INTEGER NOT NULL DEFAULT 0,
    note             TEXT,
    active           INTEGER NOT NULL DEFAULT 1,
    created_at       TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS quotes (
    id             INTEGER PRIMARY KEY AUTOINCREMENT,
    quote_no       TEXT UNIQUE NOT NULL,
    student_id     INTEGER,
    student_name   TEXT NOT NULL,
    issue_date     TEXT NOT NULL,
    note           TEXT,
    sum_tuition    INTEGER NOT NULL DEFAULT 0,
    sum_material   INTEGER NOT NULL DEFAULT 0,
    sum_deduction  INTEGER NOT NULL DEFAULT 0,
    sum_total      INTEGER NOT NULL DEFAULT 0,
    payment_status TEXT NOT NULL DEFAULT '未繳',
    payment_method TEXT,
    receipt_no     TEXT,
    created_at     TEXT NOT NULL,
    FOREIGN KEY (student_id) REFERENCES students(id) ON DELETE SET NULL
);

CREATE TABLE IF NOT EXISTS quote_items (
    id         INTEGER PRIMARY KEY AUTOINCREMENT,
    quote_id   INTEGER NOT NULL,
    name       TEXT NOT NULL,
    date_range TEXT,
    tuition    INTEGER NOT NULL DEFAULT 0,
    material   INTEGER NOT NULL DEFAULT 0,
    deduction  INTEGER NOT NULL DEFAULT 0,
    total      INTEGER NOT NULL DEFAULT 0,
    sort_order INTEGER NOT NULL DEFAULT 0,
    FOREIGN KEY (quote_id) REFERENCES quotes(id) ON DELETE CASCADE
);
"""


@contextmanager
def connect():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    try:
        yield conn
        conn.commit()
    finally:
        conn.close()


def init_db():
    with connect() as conn:
        conn.executescript(SCHEMA)


def _now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


# ─────────────────────────── 學生 ───────────────────────────
def add_student(**f) -> int:
    with connect() as conn:
        cur = conn.execute(
            """INSERT INTO students (name, grade, parent_name, phone, email, address, note, created_at)
               VALUES (?,?,?,?,?,?,?,?)""",
            (f.get("name"), f.get("grade"), f.get("parent_name"), f.get("phone"),
             f.get("email"), f.get("address"), f.get("note"), _now()),
        )
        return cur.lastrowid


def update_student(sid: int, **f):
    with connect() as conn:
        conn.execute(
            """UPDATE students SET name=?, grade=?, parent_name=?, phone=?,
               email=?, address=?, note=? WHERE id=?""",
            (f.get("name"), f.get("grade"), f.get("parent_name"), f.get("phone"),
             f.get("email"), f.get("address"), f.get("note"), sid),
        )


def delete_student(sid: int):
    with connect() as conn:
        conn.execute("DELETE FROM students WHERE id=?", (sid,))


def list_students() -> list[dict]:
    with connect() as conn:
        return [dict(r) for r in conn.execute(
            "SELECT * FROM students ORDER BY name")]


def get_student(sid: int) -> dict | None:
    with connect() as conn:
        r = conn.execute("SELECT * FROM students WHERE id=?", (sid,)).fetchone()
        return dict(r) if r else None


# ─────────────────────────── 課程／價目 ───────────────────────────
def add_course(**f) -> int:
    with connect() as conn:
        cur = conn.execute(
            """INSERT INTO courses (name, default_tuition, default_material, note, active, created_at)
               VALUES (?,?,?,?,?,?)""",
            (f.get("name"), int(f.get("default_tuition") or 0),
             int(f.get("default_material") or 0), f.get("note"),
             int(f.get("active", 1)), _now()),
        )
        return cur.lastrowid


def update_course(cid: int, **f):
    with connect() as conn:
        conn.execute(
            """UPDATE courses SET name=?, default_tuition=?, default_material=?,
               note=?, active=? WHERE id=?""",
            (f.get("name"), int(f.get("default_tuition") or 0),
             int(f.get("default_material") or 0), f.get("note"),
             int(f.get("active", 1)), cid),
        )


def delete_course(cid: int):
    with connect() as conn:
        conn.execute("DELETE FROM courses WHERE id=?", (cid,))


def list_courses(only_active: bool = False) -> list[dict]:
    q = "SELECT * FROM courses"
    if only_active:
        q += " WHERE active=1"
    q += " ORDER BY name"
    with connect() as conn:
        return [dict(r) for r in conn.execute(q)]


def get_course(cid: int) -> dict | None:
    with connect() as conn:
        r = conn.execute("SELECT * FROM courses WHERE id=?", (cid,)).fetchone()
        return dict(r) if r else None


# ─────────────────────────── 報價 ───────────────────────────
def _gen_quote_no(conn) -> str:
    d = datetime.now()
    prefix = f"QT-{d.strftime('%Y%m%d')}"
    n = conn.execute(
        "SELECT COUNT(*) FROM quotes WHERE quote_no LIKE ?", (prefix + "%",)
    ).fetchone()[0]
    return f"{prefix}-{n + 1:03d}"


def create_quote(student_id, student_name, issue_date, note, items,
                 payment_status="未繳", payment_method=None, receipt_no=None) -> int:
    """items: list of dict(name, date_range, tuition, material, deduction)。
    自動計算每列 total 與各項小計，回傳 quote id。"""
    st = sm = sd = stot = 0
    norm = []
    for it in items:
        tu = int(it.get("tuition") or 0)
        ma = int(it.get("material") or 0)
        de = int(it.get("deduction") or 0)
        tot = tu + ma - de
        st += tu; sm += ma; sd += de; stot += tot
        norm.append((it.get("name", ""), it.get("date_range", ""), tu, ma, de, tot))

    with connect() as conn:
        quote_no = _gen_quote_no(conn)
        cur = conn.execute(
            """INSERT INTO quotes (quote_no, student_id, student_name, issue_date, note,
               sum_tuition, sum_material, sum_deduction, sum_total,
               payment_status, payment_method, receipt_no, created_at)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (quote_no, student_id, student_name, issue_date, note,
             st, sm, sd, stot, payment_status, payment_method, receipt_no, _now()),
        )
        qid = cur.lastrowid
        for i, (nm, dr, tu, ma, de, tot) in enumerate(norm):
            conn.execute(
                """INSERT INTO quote_items (quote_id, name, date_range, tuition,
                   material, deduction, total, sort_order) VALUES (?,?,?,?,?,?,?,?)""",
                (qid, nm, dr, tu, ma, de, tot, i),
            )
        return qid


def list_quotes() -> list[dict]:
    with connect() as conn:
        return [dict(r) for r in conn.execute(
            "SELECT * FROM quotes ORDER BY id DESC")]


def get_quote(qid: int) -> dict | None:
    """回傳含 items 的完整報價，可直接餵給 quote_docx.render。"""
    with connect() as conn:
        q = conn.execute("SELECT * FROM quotes WHERE id=?", (qid,)).fetchone()
        if not q:
            return None
        q = dict(q)
        q["items"] = [dict(r) for r in conn.execute(
            "SELECT * FROM quote_items WHERE quote_id=? ORDER BY sort_order", (qid,))]
        return q


def update_payment(qid: int, payment_status, payment_method=None, receipt_no=None):
    with connect() as conn:
        conn.execute(
            "UPDATE quotes SET payment_status=?, payment_method=?, receipt_no=? WHERE id=?",
            (payment_status, payment_method, receipt_no, qid),
        )


def delete_quote(qid: int):
    with connect() as conn:
        conn.execute("DELETE FROM quotes WHERE id=?", (qid,))


def quote_to_docx_ctx(qid: int) -> dict:
    """把資料庫中的報價轉成 quote_docx.render 需要的格式。"""
    q = get_quote(qid)
    if not q:
        raise ValueError(f"找不到報價 id={qid}")
    return {
        "student_name": q["student_name"],
        "note": q["note"] or "",
        "courses": [
            {"name": it["name"], "date": it["date_range"],
             "tuition": it["tuition"], "material": it["material"],
             "deduction": it["deduction"]}
            for it in q["items"]
        ],
    }


if __name__ == "__main__":
    init_db()
    print(f"資料庫已初始化：{DB_PATH}")
