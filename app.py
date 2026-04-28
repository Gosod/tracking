# -*- coding: utf-8 -*-
import os
import time
import threading
from flask import Flask, request, jsonify, render_template, redirect, url_for

from db import (
    init_db, get_all_orders, get_order_by_number, create_order,
    delete_order, get_positions, insert_positions, get_position,
    add_marking, get_unsynced_markings, mark_synced, get_done_qty,
    get_conn
)
from excel_parser import parse_excel
import sheets

UPLOAD_FOLDER = os.path.join(os.path.dirname(__file__), "uploads")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

app = Flask(__name__)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024

init_db()


def sync_to_sheets():
    """
    Sync strategy: group unsynced markings by position,
    send total done per position (upsert). One API call per position.
    """
    if not sheets.is_configured():
        return

    unsynced = get_unsynced_markings()
    if not unsynced:
        return

    # Group by position_id
    positions_to_sync = {}
    for m in unsynced:
        pid = m["position_id"]
        if pid not in positions_to_sync:
            positions_to_sync[pid] = m  # keep metadata from first record

    for pid, m in positions_to_sync.items():
        try:
            total_done = get_done_qty(pid)
            sheets.upsert_position(
                order_number=m["order_number"],
                designation=m["designation"],
                name=m["name"],
                total_qty_done=total_done
            )
            # Mark all markings for this position as synced
            conn = get_conn()
            conn.execute(
                "UPDATE markings SET synced = 1 WHERE position_id = ? AND synced = 0",
                (pid,)
            )
            conn.commit()
            conn.close()
            time.sleep(1)
        except Exception as e:
            app.logger.error(f"Sync error for position {pid}: {e}")
            break


def sync_background():
    t = threading.Thread(target=sync_to_sheets, daemon=True)
    t.start()


@app.route("/")
def index():
    orders = get_all_orders()
    return render_template("index.html", orders=orders,
                           sheets_ok=sheets.is_configured())


@app.route("/order/<int:order_id>")
def order_view(order_id):
    positions = get_positions(order_id)
    if not positions:
        return redirect(url_for("index"))

    pos_list = []
    for p in positions:
        done = get_done_qty(p["id"])
        pos_list.append({
            "id": p["id"],
            "pos_number": p["pos_number"],
            "designation": p["designation"],
            "name": p["name"],
            "qty": p["qty"],
            "done": done,
            "complete": done >= p["qty"] and p["qty"] > 0
        })

    conn = get_conn()
    order = conn.execute("SELECT * FROM orders WHERE id = ?", (order_id,)).fetchone()
    conn.close()

    return render_template("order.html", order=order, positions=pos_list)


@app.route("/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        return jsonify({"error": "File not selected"}), 400

    file = request.files["file"]
    if not file.filename:
        return jsonify({"error": "File not selected"}), 400

    if not file.filename.endswith((".xlsx", ".xls")):
        return jsonify({"error": "Need .xlsx or .xls file"}), 400

    filepath = os.path.join(app.config["UPLOAD_FOLDER"], file.filename)
    file.save(filepath)

    try:
        result = parse_excel(filepath)
    except ValueError as e:
        os.remove(filepath)
        return jsonify({"error": str(e)}), 422

    order_number = result["order_number"]
    positions = result["positions"]

    existing = get_order_by_number(order_number)
    if existing:
        os.remove(filepath)
        return jsonify({
            "error": f"Order {order_number} already loaded",
            "order_id": existing["id"]
        }), 409

    order = create_order(order_number)
    insert_positions(order["id"], positions)
    os.remove(filepath)

    return jsonify({
        "ok": True,
        "order_id": order["id"],
        "order_number": order_number,
        "positions_count": len(positions)
    })


@app.route("/mark/<int:position_id>", methods=["POST"])
def mark(position_id):
    data = request.get_json(silent=True) or {}
    qty_done = data.get("qty_done", 1)

    try:
        qty_done = float(qty_done)
    except (ValueError, TypeError):
        return jsonify({"error": "Invalid quantity"}), 400

    pos = get_position(position_id)
    if not pos:
        return jsonify({"error": "Position not found"}), 404

    add_marking(position_id, qty_done)
    total_done = get_done_qty(position_id)

    sync_background()

    return jsonify({
        "ok": True,
        "total_done": total_done,
        "qty": pos["qty"],
        "complete": total_done >= pos["qty"] and pos["qty"] > 0
    })


@app.route("/delete/<int:order_id>", methods=["POST"])
def delete(order_id):
    delete_order(order_id)
    return jsonify({"ok": True})


@app.route("/sync")
def manual_sync():
    sync_to_sheets()
    return jsonify({"ok": True})


@app.route("/api/orders")
def api_orders():
    orders = get_all_orders()
    return jsonify([dict(o) for o in orders])


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)