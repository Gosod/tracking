"""
db.py — работа с локальной SQLite базой.
Хранит заказы и позиции. Буфер на случай потери WiFi.
"""

import sqlite3
import os

DB_PATH = os.path.join(os.path.dirname(__file__), "tracker.db")


def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    """Создать таблицы если не существуют."""
    conn = get_conn()
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS orders (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            order_number TEXT NOT NULL UNIQUE,  -- номер заказа из 1С (A1 Excel)
            created_at  TEXT DEFAULT (datetime('now', 'localtime'))
        );

        CREATE TABLE IF NOT EXISTS positions (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            order_id    INTEGER NOT NULL REFERENCES orders(id),
            pos_number  INTEGER,                -- Поз. (колонка 1)
            designation TEXT,                   -- Обозначение (колонка 4)
            name        TEXT,                   -- Наименование (колонка 5)
            qty         REAL DEFAULT 0          -- Кол-во (колонка 6)
        );

        CREATE TABLE IF NOT EXISTS markings (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            position_id INTEGER NOT NULL REFERENCES positions(id),
            qty_done    REAL NOT NULL,           -- сколько штук отметили готовыми
            marked_at   TEXT DEFAULT (datetime('now', 'localtime')),
            synced      INTEGER DEFAULT 0        -- 0 = не отправлено в Sheets, 1 = отправлено
        );
    """)
    conn.commit()
    conn.close()


# ── Заказы ──────────────────────────────────────────────────────────────────

def get_all_orders():
    conn = get_conn()
    rows = conn.execute(
        "SELECT * FROM orders ORDER BY created_at DESC"
    ).fetchall()
    conn.close()
    return rows


def get_order_by_number(order_number):
    conn = get_conn()
    row = conn.execute(
        "SELECT * FROM orders WHERE order_number = ?", (order_number,)
    ).fetchone()
    conn.close()
    return row


def create_order(order_number):
    conn = get_conn()
    try:
        conn.execute(
            "INSERT INTO orders (order_number) VALUES (?)", (order_number,)
        )
        conn.commit()
        row = conn.execute(
            "SELECT * FROM orders WHERE order_number = ?", (order_number,)
        ).fetchone()
        return row
    finally:
        conn.close()


def delete_order(order_id):
    """Удалить заказ и все его позиции (каскадно)."""
    conn = get_conn()
    conn.execute("DELETE FROM markings WHERE position_id IN (SELECT id FROM positions WHERE order_id = ?)", (order_id,))
    conn.execute("DELETE FROM positions WHERE order_id = ?", (order_id,))
    conn.execute("DELETE FROM orders WHERE id = ?", (order_id,))
    conn.commit()
    conn.close()


# ── Позиции ──────────────────────────────────────────────────────────────────

def get_positions(order_id):
    conn = get_conn()
    rows = conn.execute(
        "SELECT * FROM positions WHERE order_id = ? ORDER BY pos_number", (order_id,)
    ).fetchall()
    conn.close()
    return rows


def insert_positions(order_id, positions):
    """
    positions — список dict: {pos_number, designation, name, qty}
    """
    conn = get_conn()
    conn.executemany(
        "INSERT INTO positions (order_id, pos_number, designation, name, qty) VALUES (:order_id, :pos_number, :designation, :name, :qty)",
        [{"order_id": order_id, **p} for p in positions]
    )
    conn.commit()
    conn.close()


def get_position(position_id):
    conn = get_conn()
    row = conn.execute("SELECT * FROM positions WHERE id = ?", (position_id,)).fetchone()
    conn.close()
    return row


# ── Отметки готовности ────────────────────────────────────────────────────────

def add_marking(position_id, qty_done):
    conn = get_conn()
    conn.execute(
        "INSERT INTO markings (position_id, qty_done) VALUES (?, ?)",
        (position_id, qty_done)
    )
    conn.commit()
    # возвращаем запись
    row = conn.execute(
        "SELECT * FROM markings WHERE position_id = ? ORDER BY id DESC LIMIT 1",
        (position_id,)
    ).fetchone()
    conn.close()
    return row


def get_unsynced_markings():
    """Все отметки, ещё не отправленные в Google Sheets."""
    conn = get_conn()
    rows = conn.execute("""
        SELECT m.*, p.designation, p.name, p.qty,
               o.order_number
        FROM markings m
        JOIN positions p ON m.position_id = p.id
        JOIN orders o ON p.order_id = o.id
        WHERE m.synced = 0
    """).fetchall()
    conn.close()
    return rows


def mark_synced(marking_id):
    conn = get_conn()
    conn.execute("UPDATE markings SET synced = 1 WHERE id = ?", (marking_id,))
    conn.commit()
    conn.close()


def get_done_qty(position_id):
    """Суммарное кол-во готовых штук по позиции."""
    conn = get_conn()
    row = conn.execute(
        "SELECT COALESCE(SUM(qty_done), 0) as total FROM markings WHERE position_id = ?",
        (position_id,)
    ).fetchone()
    conn.close()
    return row["total"] if row else 0
